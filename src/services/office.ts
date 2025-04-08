import html2canvas from 'html2canvas';

export const OfficeService = {
  async insertContent(content: any) {
    try {
      // Check if running in Office context
      if (!window.Office) {
        console.error('Office JS API not available');
        return;
      }

      // Format the content as HTML
      const htmlContent = this.formatContentAsHtml(content);

      // Insert the formatted content into the document
      return new Promise((resolve, reject) => {
        Office.context.document.setSelectedDataAsync(
          htmlContent,
          { coercionType: Office.CoercionType.Html },
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
      
      // First capture the preview as an image
      const canvas = await html2canvas(element, {
        logging: true,
        useCORS: true,
        allowTaint: true,
        background: '#ffffff',
        width: element.offsetWidth,
        height: element.offsetHeight
      });
      
      console.log('Preview captured as canvas');
      
      // Convert to data URL with high quality
      const imageDataUrl = canvas.toDataURL('image/png', 1.0);
      console.log('Canvas converted to data URL');

      // Create HTML with embedded image and styling
      const htmlContent = `
        <div style="
          font-family: 'Segoe UI', sans-serif;
          margin: 10px 0;
          max-width: 100%;
        ">
          <p style="
            color: #666;
            font-size: 11px;
            margin: 0 0 8px 0;
          ">Dashboard Preview:</p>
          <div style="
            border: 1px solid #e1e1e1;
            border-radius: 6px;
            padding: 2px;
            background: white;
            box-shadow: 0 2px 4px rgba(0,0,0,0.05);
          ">
            <img 
              src="${imageDataUrl}" 
              alt="Dashboard Preview" 
              style="
                display: block;
                width: 100%;
                max-width: 600px;
                height: auto;
                border-radius: 4px;
              "
            />
          </div>
        </div>
      `;

      return new Promise((resolve, reject) => {
        try {
          // Insert the HTML with embedded image
          Office.context.document.setSelectedDataAsync(
            htmlContent,
            { coercionType: Office.CoercionType.Html },
            (result) => {
              if (result.status === Office.AsyncResultStatus.Succeeded) {
                console.log('Preview inserted successfully');
                resolve();
              } else {
                console.error('Failed to insert preview:', result.error);
                reject(new Error('Failed to insert preview'));
              }
            }
          );
        } catch (error) {
          console.error('Error in preview insertion:', error);
          reject(error);
        }
      });
    } catch (error) {
      console.error('Error capturing preview:', error);
      throw error;
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

  formatContentAsHtml(content: any): string {
    const isMetric = content && content.hasOwnProperty('id') && content.hasOwnProperty('label');
    
    if (isMetric) {
      return `
        <div style="font-family: 'Segoe UI', sans-serif; padding: 15px; border: 1px solid #e1e1e1; border-radius: 6px; max-width: 600px;">
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
          <div style="font-size: 11px; color: #a19f9d; border-top: 1px solid #e1e1e1; padding-top: 8px;">
            Metric ID: ${content.id || 'Unknown'}
          </div>
        </div>
      `;
    } else {
      return `
        <div style="font-family: 'Segoe UI', sans-serif; padding: 15px; border: 1px solid #e1e1e1; border-radius: 6px; max-width: 600px;">
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
          <div style="font-size: 11px; color: #a19f9d; border-top: 1px solid #e1e1e1; padding-top: 8px;">
            Dashboard ID: ${content.Id || 'Unknown'}
          </div>
        </div>
      `;
    }
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
  }
}; 