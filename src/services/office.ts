export const OfficeService = {
  async insertContent(content: any) {
    try {
      // Check if running in Office context
      if (!window.Office) {
        console.error('Office JS API not available');
        return;
      }

      // Depending on the type of content, format it accordingly
      const isMetric = content && content.hasOwnProperty('id') && content.hasOwnProperty('label');
      let htmlContent = '';

      if (isMetric) {
        // Format metric data
        htmlContent = `
          <div style="font-family: 'Segoe UI', sans-serif; padding: 15px; border: 1px solid #e1e1e1; border-radius: 6px; max-width: 600px;">
            <div style="display: flex; align-items: center; margin-bottom: 12px;">
              <div style="font-size: 18px; font-weight: 600; color: #0078d4;">${content.label || 'Unknown Metric'}</div>
              <div style="margin-left: auto; font-size: 12px; color: #605e5c;">${content.type || 'Metric'}</div>
            </div>
            ${content.description ? `<div style="margin-bottom: 12px;">${content.description}</div>` : ''}
            <table style="width: 100%; border-collapse: collapse; font-size: 13px; margin-bottom: 12px;">
              <tr>
                <td style="padding: 4px 0; color: #605e5c;">Creator:</td>
                <td style="padding: 4px 0; text-align: right;">${content.creatorName || 'Unknown'}</td>
              </tr>
              <tr>
                <td style="padding: 4px 0; color: #605e5c;">Created:</td>
                <td style="padding: 4px 0; text-align: right;">${formatDate(content.createdDate)}</td>
              </tr>
              <tr>
                <td style="padding: 4px 0; color: #605e5c;">Workspace:</td>
                <td style="padding: 4px 0; text-align: right;">${content.namespace || 'Default'}</td>
              </tr>
            </table>
            <div style="font-size: 11px; color: #a19f9d; border-top: 1px solid #e1e1e1; padding-top: 8px;">
              Metric ID: ${content.id || 'Unknown'}
            </div>
          </div>
        `;
      } else {
        // Format dashboard data
        htmlContent = `
          <div style="font-family: 'Segoe UI', sans-serif; padding: 15px; border: 1px solid #e1e1e1; border-radius: 6px; max-width: 600px;">
            <div style="display: flex; align-items: center; margin-bottom: 12px;">
              <div style="font-size: 18px; font-weight: 600; color: #0078d4;">${content.MasterLabel || content.Name || 'Unknown Dashboard'}</div>
              <div style="margin-left: auto; font-size: 12px; color: #605e5c;">${content.AnalyticsWorkspace?.MasterLabel || 'Dashboard'}</div>
            </div>
            ${content.Description ? `<div style="margin-bottom: 12px;">${content.Description}</div>` : ''}
            <table style="width: 100%; border-collapse: collapse; font-size: 13px; margin-bottom: 12px;">
              <tr>
                <td style="padding: 4px 0; color: #605e5c;">Creator:</td>
                <td style="padding: 4px 0; text-align: right;">${content.CreatedBy?.Name || 'Unknown'}</td>
              </tr>
              <tr>
                <td style="padding: 4px 0; color: #605e5c;">Created:</td>
                <td style="padding: 4px 0; text-align: right;">${formatDate(content.CreatedDate)}</td>
              </tr>
              <tr>
                <td style="padding: 4px 0; color: #605e5c;">Workspace:</td>
                <td style="padding: 4px 0; text-align: right;">${content.AnalyticsWorkspace?.MasterLabel || 'Default'}</td>
              </tr>
            </table>
            <div style="font-size: 11px; color: #a19f9d; border-top: 1px solid #e1e1e1; padding-top: 8px;">
              Dashboard ID: ${content.Id || 'Unknown'}
            </div>
          </div>
        `;
      }

      // Insert the formatted content into the document
      await window.Office.context.document.setSelectedDataAsync(
        htmlContent,
        {
          coercionType: window.Office.CoercionType.Html
        },
        (result) => {
          if (result.status === window.Office.AsyncResultStatus.Succeeded) {
            console.log('Content inserted successfully');
          } else {
            console.error('Error inserting content:', result.error);
          }
        }
      );

      console.log('Content insertion process completed');
      return true;
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
  }
};

// Helper function to format dates
function formatDate(dateString: string): string {
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