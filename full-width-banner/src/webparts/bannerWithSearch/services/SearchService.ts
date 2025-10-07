import { WebPartContext } from '@microsoft/sp-webpart-base';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { ISearchResult } from '../components/IBannerWithSearchProps';

export class SearchService {
  private context: WebPartContext;

  constructor(context: WebPartContext) {
    this.context = context;
  }

  // Combined adhoc suggestions for files and people
  public async searchAdhocCombined(query: string, limitToSite: boolean): Promise<ISearchResult[]> {
    const docs = await this.searchSharePoint(query, 'Documents', limitToSite);
    const people = await this.searchPeople(query, limitToSite);
    return [...people, ...docs];
  }

  private async searchPeople(query: string, limitToSite: boolean): Promise<ISearchResult[]> {
    const scope = limitToSite ? `&SiteUrl=${encodeURIComponent(this.context.pageContext.web.absoluteUrl)}` : '';
    const selectProps = 'Title,Path,WorkEmail,Department,JobTitle,OfficeNumber,PreferredName,PictureURL';
    const url = `${this.context.pageContext.web.absoluteUrl}/_api/search/query?querytext='${encodeURIComponent(query)}'&rowlimit=5${scope}&selectproperties='${selectProps}'&sourceid='b09a7990-05ea-4af9-81ef-edfab16c4e31'`;
    const response: SPHttpClientResponse = await this.context.spHttpClient.get(url, SPHttpClient.configurations.v1);
    if (!response.ok) return [];
    const data = await response.json();
    const results: ISearchResult[] = [];
    if (data.PrimaryQueryResult && data.PrimaryQueryResult.RelevantResults) {
      data.PrimaryQueryResult.RelevantResults.Table.Rows.forEach((row: { Cells: { Key: string; Value: string }[] }) => {
        const cells = row.Cells;
        results.push({
          title: this.getCellValue(cells, 'PreferredName') || this.getCellValue(cells, 'Title') || '',
          description: this.getCellValue(cells, 'WorkEmail') || '',
          author: '',
          modifiedDate: '',
          url: this.getCellValue(cells, 'Path') || '',
          userImage: this.getCellValue(cells, 'PictureURL') || '',
          department: this.getCellValue(cells, 'Department') || '',
          jobTitle: this.getCellValue(cells, 'JobTitle') || '',
          officeLocation: this.getCellValue(cells, 'OfficeNumber') || '',
          status: ''
        });
      });
    }
    return results;
  }

  public async search(query: string, resultType: string, category: 'Documents' | 'People' | 'Sites', limitToSite: boolean): Promise<ISearchResult[]> {
    try {
      if (resultType === 'SharePoint') {
        return await this.searchSharePoint(query, category, limitToSite);
      } else {
        return await this.searchAdhoc(query);
      }
    } catch (error) {
      console.error('Search service error:', error);
      return [];
    }
  }

  private async searchSharePoint(query: string, category: 'Documents' | 'People' | 'Sites', limitToSite: boolean): Promise<ISearchResult[]> {
    let queryText = encodeURIComponent(query);
    let selectProps = 'Title,Path,Author,LastModifiedTime,FileType,FileSize,Description';
    let extraParams = '';

    if (category === 'People') {
      // People search: use the built-in Local People Results source
      selectProps = 'Title,Path,WorkEmail,Department,JobTitle,OfficeNumber,AccountName,PreferredName,PictureURL';
      extraParams += `&sourceid='b09a7990-05ea-4af9-81ef-edfab16c4e31'`;
    } else if (category === 'Sites') {
      // Restrict to sites
      queryText = encodeURIComponent(`${query} (contentclass:STS_Site OR contentclass:STS_Web)`);
    } else {
      // Documents: optionally bias towards files
      // No special source; keep defaults
    }

    const scope = limitToSite ? `&SiteUrl=${encodeURIComponent(this.context.pageContext.web.absoluteUrl)}` : '';

    const searchUrl = `${this.context.pageContext.web.absoluteUrl}/_api/search/query?querytext='${queryText}'&rowlimit=10${scope}&selectproperties='${selectProps}'${extraParams}`;
    
    const response: SPHttpClientResponse = await this.context.spHttpClient.get(
      searchUrl,
      SPHttpClient.configurations.v1
    );

    if (!response.ok) {
      throw new Error(`Search failed: ${response.status}`);
    }

    const data = await response.json();
    const results: ISearchResult[] = [];

    if (data.PrimaryQueryResult && data.PrimaryQueryResult.RelevantResults) {
      data.PrimaryQueryResult.RelevantResults.Table.Rows.forEach((row: { Cells: { Key: string; Value: string }[] }) => {
        const cells = row.Cells;
        const result: ISearchResult = {
          title: this.getCellValue(cells, 'Title') || 'Untitled',
          description: this.getCellValue(cells, 'Description') || '',
          author: this.getCellValue(cells, 'Author') || 'Unknown',
          modifiedDate: this.formatDate(this.getCellValue(cells, 'LastModifiedTime')),
          url: this.getCellValue(cells, 'Path') || '',
          fileType: this.getCellValue(cells, 'FileType'),
          fileSize: this.formatFileSize(parseInt(this.getCellValue(cells, 'FileSize') || '0')),
          icon: this.getFileIcon(this.getCellValue(cells, 'FileType'))
        };
        results.push(result);
      });
    }

    return results;
  }

  private async searchAdhoc(query: string): Promise<ISearchResult[]> {
    // Mock data for adhoc search - in real implementation, this would call your custom search API
    const mockResults: ISearchResult[] = [
      {
        title: 'John Doe',
        description: 'Senior Software Engineer',
        author: 'John Doe',
        modifiedDate: '2024-01-15',
        url: `${this.context.pageContext.web.absoluteUrl}/_layouts/15/userdisp.aspx?ID=1`,
        userImage: `${this.context.pageContext.web.absoluteUrl}/_layouts/15/userphoto.aspx?size=L&username=john.doe`,
        department: 'Engineering',
        officeLocation: 'Seattle, WA',
        jobTitle: 'Senior Software Engineer',
        status: 'Active'
      },
      {
        title: 'Jane Smith',
        description: 'Product Manager',
        author: 'Jane Smith',
        modifiedDate: '2024-01-14',
        url: `${this.context.pageContext.web.absoluteUrl}/_layouts/15/userdisp.aspx?ID=2`,
        userImage: `${this.context.pageContext.web.absoluteUrl}/_layouts/15/userphoto.aspx?size=L&username=jane.smith`,
        department: 'Product',
        officeLocation: 'New York, NY',
        jobTitle: 'Product Manager',
        status: 'Away'
      }
    ];

    // Filter results based on query
    return mockResults.filter(result => 
      result.title.toLowerCase().indexOf(query.toLowerCase()) !== -1 ||
      (result.department && result.department.toLowerCase().indexOf(query.toLowerCase()) !== -1) ||
      (result.jobTitle && result.jobTitle.toLowerCase().indexOf(query.toLowerCase()) !== -1)
    );
  }

  private getCellValue(cells: { Key: string; Value: string }[], key: string): string {
    for (let i = 0; i < cells.length; i++) {
      if (cells[i].Key === key) {
        return cells[i].Value || '';
      }
    }
    return '';
  }

  private formatDate(dateString: string): string {
    if (!dateString) return '';
    const date = new Date(dateString);
    return date.toLocaleDateString();
  }

  private formatFileSize(bytes: number): string {
    if (bytes === 0) return '0 Bytes';
    const k = 1024;
    const sizes = ['Bytes', 'KB', 'MB', 'GB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
  }

  private getFileIcon(fileType: string): string {
    const iconMap: { [key: string]: string } = {
      'docx': 'WordDocument',
      'doc': 'WordDocument',
      'xlsx': 'ExcelDocument',
      'xls': 'ExcelDocument',
      'pptx': 'PowerPointDocument',
      'ppt': 'PowerPointDocument',
      'pdf': 'PDF',
      'txt': 'TextDocument',
      'jpg': 'Photo2',
      'jpeg': 'Photo2',
      'png': 'Photo2',
      'gif': 'Photo2',
      'mp4': 'Video',
      'avi': 'Video',
      'zip': 'ZipFolder',
      'rar': 'ZipFolder'
    };

    return iconMap[fileType?.toLowerCase()] || 'Page';
  }
}
