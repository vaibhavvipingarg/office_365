import html2canvas from 'html2canvas';

// Add type declarations at the top of the file
declare global {
  interface Window {
    $Lightning: {
      use: (
        appName: string,
        callback: () => void,
        orgUrl: string,
        accessToken: string,
        config?: {
          allowedDomains?: string[];
          useAppHost?: boolean;
        }
      ) => void;
      createComponent: (
        componentName: string,
        componentAttributes: any,
        containerId: string,
        callback: (cmp: any) => void
      ) => void;
    }
    Office: {
      context: {
        host: string;
        mailbox?: {
          item: {
            body: {
              setSelectedDataAsync: (
                data: string,
                options: { coercionType: string },
                callback: (result: { status: string; error?: any }) => void
              ) => void;
            };
          };
        };
        document: {
          setSelectedDataAsync: (
            data: string,
            options: { coercionType: string },
            callback: (result: { status: string; error?: any }) => void
          ) => void;
        };
      };
      AsyncResultStatus: {
        Succeeded: string;
        Failed: string;
      };
      CoercionType: {
        Text: string;
        Html: string;
        Image: string;
      };
    };
  }
}

export const OfficeService = {
  // Get the current Office host application
  getHostType(): string {
    if (!window.Office) {
      return 'unknown';
    }
    return Office.context.host;
  },

  // Check if we're running in a specific host
  isHostType(hostType: string): boolean {
    return this.getHostType().toLowerCase() === hostType.toLowerCase();
  },

  async insertContent(content: any) {
    try {
      // Check if running in Office context
      if (!window.Office) {
        console.error('Office JS API not available');
        return;
      }

      // Format the content based on the host application
      const hostType = this.getHostType();
      let formattedContent;
      let coercionType = Office.CoercionType.Html;

      switch (hostType) {
        case 'Word':
          formattedContent = this.formatContentAsHtml(content);
          break;
        case 'PowerPoint':
          formattedContent = this.formatContentForPowerPoint(content);
          coercionType = Office.CoercionType.Text; // PowerPoint supports text insertion
          break;
        case 'Outlook':
          formattedContent = this.formatContentForOutlook(content);
          break;
        default:
          formattedContent = this.formatContentAsHtml(content);
      }

      // Insert the formatted content into the document
      return new Promise((resolve, reject) => {
        if (this.isHostType('Outlook')) {
          // For Outlook, we use the mailbox item API
          Office.context.mailbox.item.body.setSelectedDataAsync(
            formattedContent,
            { coercionType: Office.CoercionType.Html },
            (result) => {
              if (result.status === Office.AsyncResultStatus.Succeeded) {
                console.log('Content inserted successfully in Outlook');
                resolve(true);
              } else {
                console.error('Error inserting content in Outlook:', result.error);
                reject(result.error);
              }
            }
          );
        } else {
          // For Word and PowerPoint, we use the document API
          Office.context.document.setSelectedDataAsync(
            formattedContent,
            { coercionType },
            (result) => {
              if (result.status === Office.AsyncResultStatus.Succeeded) {
                console.log('Content inserted successfully');
                resolve(true);
              } else {
                console.error('Error inserting content:', result.error);
                reject(result.error);
              }
            }
          );
        }
      });
    } catch (error) {
      console.error('Error inserting content into document:', error);
      throw error;
    }
  },

  async insertImage(imageUrl: string): Promise<void> {
    // Check if Office is available
    if (!window.hasOwnProperty('Office')) {
      console.log('Office.js not available - running in local mode');
      return Promise.resolve();
    }
    
    return new Promise((resolve, reject) => {
      try {
        Office.context.document.setSelectedDataAsync(
          imageUrl,
          { coercionType: Office.CoercionType.Image },
          (result) => {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
              resolve();
            } else {
              reject(new Error('Failed to insert image'));
            }
          }
        );
      } catch (error) {
        reject(error);
      }
    });
  },

  async insertHtml(html: string): Promise<void> {
    // Check if Office is available
    if (!window.hasOwnProperty('Office')) {
      console.log('Office.js not available - running in local mode');
      return Promise.resolve();
    }
    
    return new Promise((resolve, reject) => {
      try {
        Office.context.document.setSelectedDataAsync(
          html,
          { coercionType: Office.CoercionType.Html },
          (result) => {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
              resolve();
            } else {
              reject(new Error('Failed to insert HTML'));
            }
          }
        );
      } catch (error) {
        reject(error);
      }
    });
  },

  async insertImageFromElement(element: HTMLElement): Promise<void> {
    try {
      console.log('Starting preview capture process...');
      
      // Find the Lightning component container
      const lightningContainer = element.querySelector('[id^="preview-lightning-"]');
      if (!lightningContainer) {
        console.warn('Lightning container not found, falling back to full preview');
        await this.insertAsHtml(element);
        return;
      }

      try {
        console.log('Attempting to capture Lightning component as image');
        const canvas = await html2canvas(lightningContainer as HTMLElement, {
          logging: true,
          useCORS: true,
          allowTaint: true,
          background: '#ffffff'
        });
        
        console.log('Lightning component captured as canvas');
        
        const imageDataUrl = canvas.toDataURL('image/png', 1.0);
        console.log('Canvas converted to data URL');

        // Get the dashboard title from the preview card
        const titleElement = element.querySelector('.preview-card .ms-Text');
        const title = titleElement?.textContent || 'Dashboard';

        const htmlContent = `
          <div style="font-family: 'Segoe UI', sans-serif; margin: 10px 0;">
            <div style="border: 1px solid #e1e1e1; border-radius: 6px; padding: 15px; background: white;">
              <div style="font-size: 18px; font-weight: 600; color: #0078d4; margin-bottom: 10px;">
                ${title}
              </div>
              <img 
                src="${imageDataUrl}" 
                alt="${title}" 
                style="display: block; width: 100%; max-width: 800px; height: auto; margin: 0 auto;"
              />
            </div>
          </div>
        `;

        await this.insertHtml(htmlContent);
        console.log('Lightning component image inserted successfully');
      } catch (captureError) {
        console.error('Failed to capture Lightning component:', captureError);
        console.warn('Falling back to full HTML preview');
        await this.insertAsHtml(element);
      }
    } catch (error) {
      console.error('Error inserting content:', error);
      throw error;
    }
  },

  initializeLightningComponent(containerId: string, dashboardId: string, accessToken: string): Promise<void> {
    console.log('Initializing Lightning component...', { containerId, dashboardId });
    
    return new Promise((resolve, reject) => {
      try {
        // Load Lightning Out script if not already loaded
        if (!document.querySelector('script[src*="lightning.out.js"]')) {
          const script = document.createElement('script');
          // Add timestamp to prevent caching
          script.src = `https://sdb42com6.test13.my.pc-rnd.salesforce.com/lightning/lightning.out.js?_=${Date.now()}`;
          script.onload = () => {
            this.createLightningComponent(containerId, dashboardId, accessToken, resolve, reject);
          };
          script.onerror = (error) => {
            console.error('Failed to load Lightning Out script:', error);
            reject(error);
          };
          document.head.appendChild(script);
        } else {
          this.createLightningComponent(containerId, dashboardId, accessToken, resolve, reject);
        }
      } catch (error) {
        console.error('Error initializing Lightning Out:', error);
        reject(error);
      }
    });
  },

  // Helper method to create the Lightning component
  createLightningComponent(
    containerId: string, 
    dashboardId: string, 
    accessToken: string,
    resolve: () => void,
    reject: (error: Error) => void
  ): void {
    if (typeof window.$Lightning !== 'undefined') {
      const targetOrigin = 'https://vaibhavvipingarg.github.io';
      console.log('Creating Lightning component with origin:', targetOrigin);

      window.$Lightning.use(
        "tableau_einstein:tableauEinsteinApp",
        () => {
          console.log('Lightning app initialized, creating component...');
          window.$Lightning.createComponent(
            "analytics_embedding:dashboard3p",
            {
              height: 300,
              idOrApiName: dashboardId,
              allowTransparency: true,
              showHeader: false,
              showSharing: false
            },
            containerId,
            (cmp: any) => {
              if (cmp) {
                console.log("Lightning component created successfully");
                // Store the component reference and mark as ready
                const container = document.getElementById(containerId);
                if (container) {
                  container.setAttribute('data-lightning-ready', 'true');
                  // Store the container ID for later use
                  container.setAttribute('data-container-id', containerId);
                  resolve();
                }
              } else {
                console.error("Failed to create Lightning component");
                reject(new Error("Failed to create Lightning component"));
              }
            }
          );
        },
        'https://sdb42com6.test13.lightning.pc-rnd.force.com',
        accessToken,
        {
          allowedDomains: [targetOrigin],
          useAppHost: true
        }
      );
    } else {
      console.error('Lightning Out not available');
      reject(new Error('Lightning Out not available'));
    }
  },

  formatContentAsHtml(content: any): string {
    // If we have a captured image, use that regardless of content type
    if (content.capturedImage) {
      const title = content.hasOwnProperty('id') 
        ? (content.label || 'Metric Card')
        : (content.MasterLabel || content.Name || 'Dashboard');

      return `
        <div style="font-family: 'Segoe UI', sans-serif; margin: 10px 0;">
          <div style="padding: 15px;">
            <div style="font-size: 18px; font-weight: 600; color: #0078d4; margin-bottom: 10px;">
              ${title}
            </div>
            <img 
              src="${content.capturedImage}" 
              alt="${title}" 
              style="display: block; width: 100%; max-width: 800px; height: auto; margin: 0 auto;"
            />
            <div style="font-size: 11px; color: #a19f9d; margin-top: 10px;">
              ${content.hasOwnProperty('id') ? `Metric ID: ${content.id}` : `Dashboard ID: ${content.Id || 'Unknown'}`}
            </div>
          </div>
        </div>
      `;
    }

    // Fallback to the old HTML format if no captured image
    const isMetric = content && content.hasOwnProperty('id') && content.hasOwnProperty('label');
    
    if (isMetric) {
      return `
        <div style="font-family: 'Segoe UI', sans-serif; padding: 15px; max-width: 600px;">
          <div style="display: flex; align-items: center; margin-bottom: 12px;">
            <div style="font-size: 18px; font-weight: 600; color: #0078d4;">${content.label || 'Unknown Metric'}</div>
            <div style="margin-left: auto; font-size: 12px; color: #605e5c;">${content.type || 'Metric'}</div>
          </div>
          ${content.description ? `<div style="margin-bottom: 12px;">${content.description}</div>` : ''}
          <div style="font-size: 13px; margin-bottom: 8px;">
            <span style="color: #605e5c;">Creator:</span>
            <span style="float: right;">${content.creatorName || 'Unknown'}</span>
          </div>
          <div style="font-size: 13px; margin-bottom: 8px;">
            <span style="color: #605e5c;">Created:</span>
            <span style="float: right;">${this.formatDate(content.createdDate)}</span>
          </div>
          <div style="font-size: 13px; margin-bottom: 8px;">
            <span style="color: #605e5c;">Workspace:</span>
            <span style="float: right;">${content.namespace || 'Default'}</span>
          </div>
          <div style="font-size: 11px; color: #a19f9d;">
            Metric ID: ${content.id || 'Unknown'}
          </div>
        </div>
      `;
    }

    return `
      <div style="font-family: 'Segoe UI', sans-serif; padding: 15px; max-width: 600px;">
        <div style="display: flex; align-items: center; margin-bottom: 12px;">
          <div style="font-size: 18px; font-weight: 600; color: #0078d4;">${content.MasterLabel || content.Name || 'Unknown Dashboard'}</div>
          <div style="margin-left: auto; font-size: 12px; color: #605e5c;">${content.AnalyticsWorkspace?.MasterLabel || 'Dashboard'}</div>
        </div>
        ${content.Description ? `<div style="margin-bottom: 12px;">${content.Description}</div>` : ''}
        <div style="font-size: 13px; margin-bottom: 8px;">
          <span style="color: #605e5c;">Creator:</span>
          <span style="float: right;">${content.CreatedBy?.Name || 'Unknown'}</span>
        </div>
        <div style="font-size: 13px; margin-bottom: 8px;">
          <span style="color: #605e5c;">Created:</span>
          <span style="float: right;">${this.formatDate(content.CreatedDate)}</span>
        </div>
        <div style="font-size: 13px; margin-bottom: 8px;">
          <span style="color: #605e5c;">Workspace:</span>
          <span style="float: right;">${content.AnalyticsWorkspace?.MasterLabel || 'Default'}</span>
        </div>
        <div style="font-size: 11px; color: #a19f9d;">
          Dashboard ID: ${content.Id || 'Unknown'}
        </div>
      </div>
    `;
  },

  formatDate(dateString: string): string {
    if (!dateString) return 'Unknown';
    
    try {
      const date = new Date(dateString);
      return date.toLocaleDateString('en-US', {
        year: 'numeric',
        month: 'short',
        day: 'numeric'
      });
    } catch (e) {
      return dateString;
    }
  },

  async insertAsHtml(element: HTMLElement): Promise<void> {
    console.log('Falling back to HTML insertion');
    const htmlContent = element.outerHTML;
    
    return new Promise((resolve, reject) => {
      Office.context.document.setSelectedDataAsync(
        htmlContent,
        { coercionType: Office.CoercionType.Html },
        (result) => {
          if (result.status === Office.AsyncResultStatus.Succeeded) {
            console.log('HTML content inserted successfully');
            resolve();
          } else {
            console.error('Failed to insert HTML content:', result.error);
            reject(new Error('Failed to insert content'));
          }
        }
      );
    });
  },

  formatContentForPowerPoint(content: any): string {
    // Format content specifically for PowerPoint as plain text
    // Since we're using text coercion, we'll create a simpler text-based layout
    const title = content.title || content.MasterLabel || content.Name || 'Salesforce Data';
    const description = content.description || content.Description || '';
    let textContent = `${title}\n\n`;

    if (description) {
      textContent += `${description}\n\n`;
    }

    // If it's a metric
    if (content.hasOwnProperty('id')) {
      textContent += `Type: ${content.type || 'Metric'}\n`;
      textContent += `Creator: ${content.creatorName || 'Unknown'}\n`;
      textContent += `Created: ${this.formatDate(content.createdDate)}\n`;
      textContent += `Workspace: ${content.namespace || 'Default'}\n`;
      textContent += `Metric ID: ${content.id || 'Unknown'}\n`;
    }
    // If it's a dashboard
    else {
      textContent += `Type: ${content.AnalyticsWorkspace?.MasterLabel || 'Dashboard'}\n`;
      textContent += `Creator: ${content.CreatedBy?.Name || 'Unknown'}\n`;
      textContent += `Created: ${this.formatDate(content.CreatedDate)}\n`;
      textContent += `Workspace: ${content.AnalyticsWorkspace?.MasterLabel || 'Default'}\n`;
      textContent += `Dashboard ID: ${content.Id || 'Unknown'}\n`;
    }

    return textContent;
  },

  formatContentForOutlook(content: any): string {
    // Format content specifically for Outlook
    // This will create an email-friendly layout
    const title = content.title || 'Salesforce Data';
    const description = content.description || '';
    const data = content.data || {};

    return `
      <div style="font-family: 'Segoe UI', sans-serif; max-width: 600px; margin: 0 auto;">
        <h3 style="color: #0078d4; border-bottom: 1px solid #ccc; padding-bottom: 10px;">${title}</h3>
        ${description ? `<p style="color: #333;">${description}</p>` : ''}
        <table style="width: 100%; border-collapse: collapse; margin-top: 15px;">
          ${Object.entries(data).map(([key, value]) => `
            <tr>
              <td style="padding: 8px; border-bottom: 1px solid #eee; font-weight: bold; width: 30%;">${key}</td>
              <td style="padding: 8px; border-bottom: 1px solid #eee;">${value}</td>
            </tr>
          `).join('')}
        </table>
      </div>
    `;
  }
}; 