// Mock implementation of Office service for standalone mode
export class OfficeService {
  static async insertContent(data: any[]): Promise<void> {
    console.log('OfficeService.insertContent called with:', data);
    return Promise.resolve();
  }

  static async insertImage(imageUrl: string): Promise<void> {
    console.log('OfficeService.insertImage called with:', imageUrl);
    return Promise.resolve();
  }

  static async insertHtml(html: string): Promise<void> {
    console.log('OfficeService.insertHtml called with:', html);
    return Promise.resolve();
  }
} 