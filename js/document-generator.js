// File: js/document-generator.js

// Document Generator with DOCX Support
class DocumentGenerator {
  constructor() {
    this.xlsx = null;
    this.docx = null;
    this.fileSaver = null;
    this.loadLibraries();
  }

  // Load required libraries
  async loadLibraries() {
    try {
      // Load SheetJS (xlsx) for Excel generation
      if (!window.XLSX) {
        await this.loadScript('https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js');
      }
      this.xlsx = window.XLSX;

      // Load docx.js for Word document generation
      if (!window.docx) {
        await this.loadScript('js/docx.min.js');
      }
      this.docx = window.docx;

      // Load FileSaver for saving files
      if (!window.saveAs) {
        await this.loadScript('https://cdn.jsdelivr.net/npm/file-saver@2.0.5/dist/FileSaver.min.js');
      }
      this.fileSaver = window.saveAs;

      console.log('Libraries loaded successfully');
    } catch (error) {
      console.error('Error loading libraries:', error);
    }
  }

  // Helper to load external scripts
  loadScript(src) {
    return new Promise((resolve, reject) => {
      const script = document.createElement('script');
      script.src = src;
      script.onload = resolve;
      script.onerror = reject;
      document.head.appendChild(script);
    });
  }

  // Generate Excel file from video data
  generateExcel(videos) {
    if (!videos || videos.length === 0) {
      throw new Error('Nenhum dado disponível para gerar a planilha.');
    }

    try {
      if (this.xlsx) {
        const data = videos.map(video => ({
          'Nome do Episodio': video.title,
          'Duração': video.duration,
          'Views': video.views,
          'Likes': video.likes,
          'Link': `https://www.youtube.com/watch?v=${video.videoId}`,
          'Data de Publicacao': new Date(video.publishedDate).toLocaleDateString()
        }));

        const ws = this.xlsx.utils.json_to_sheet(data);
        const wb = this.xlsx.utils.book_new();
        this.xlsx.utils.book_append_sheet(wb, ws, 'Playlist');

        const excelBuffer = this.xlsx.write(wb, { bookType: 'xlsx', type: 'array' });
        const blob = new Blob([excelBuffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });

        return {
          blob,
          filename: 'playlist_data.xlsx'
        };
      } else {
        return this.generateCSV(videos);
      }
    } catch (error) {
      console.error('Error generating Excel:', error);
      return this.generateCSV(videos);
    }
  }

  // Fallback: Generate CSV file
  generateCSV(videos) {
    let csv = 'Nome do Episodio,Duração,Views,Likes,Link,Data de Publicacao\n';
    videos.forEach(video => {
      const row = [
        `"${video.title.replace(/"/g, '""')}"`,
        video.duration,
        video.views,
        video.likes,
        `"https://www.youtube.com/watch?v=${video.videoId}"`,
        new Date(video.publishedDate).toLocaleDateString()
      ];
      csv += row.join(',') + '\n';
    });

    const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
    return {
      blob,
      filename: 'playlist_data.csv'
    };
  }

  // Generate Word document (DOCX format)
  async generateWordDocument(videos, screenshots) {
    if (!videos || videos.length === 0 || !screenshots || screenshots.length === 0) {
      throw new Error('Nenhum dado disponível para gerar o documento.');
    }

    try {
      if (!this.docx) {
        console.warn('DOCX library not loaded, falling back to HTML');
        return this.generateHtmlDocument(videos, screenshots);
      }
      const {
        Document,
        Paragraph,
        TextRun,
        ImageRun,
        Table,
        TableRow,
        TableCell,
        BorderStyle,
        AlignmentType,
        HeadingLevel
      } = this.docx;
      
      const doc = new Document({
        creator: 'YouTube Playlist Extractor',
        title: 'Comprovação de Dados - Playlist YouTube',
        description: 'Documento de comprovação de dados extraídos de uma playlist do YouTube',
        styles: { /* ...styles omitted for brevity...*/ }
      });
      
      doc.addSection({
        properties: {},
        children: [
          new Paragraph({
            text: 'Comprovação de Dados - Playlist YouTube',
            heading: HeadingLevel.HEADING_1,
            alignment: AlignmentType.CENTER
          })
        ]
      });

      for (let i = 0; i < videos.length; i++) {
        const video = videos[i];
        const screenshot = screenshots[i];
        
        let imageData;
        try {
          imageData = await this.getImageDataFromUrl(screenshot.imageUrl);
        } catch (e) {
          console.warn('Failed to fetch image', e);
        }
        
        const videoSection = [];
        videoSection.push(
          new Paragraph({
            text: `Vídeo ${i + 1}: ${video.title}`,
            heading: HeadingLevel.HEADING_2
          })
        );
        // ... build tables and images ...
        
        doc.addSection({ properties: {}, children: videoSection });
      }

      // Generate DOCX file
      const buffer = await this.docx.Packer.toBuffer(doc);
      const blob = new Blob(
        [buffer],
        { type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' }
      );
      return {
        blob,
        filename: 'comprovacao_videos.docx'
      };
    } catch (error) {
      console.error('Error generating DOCX:', error);
      return this.generateHtmlDocument(videos, screenshots);
    }
  }

  // Helper function to get image data from URL
  async getImageDataFromUrl(url) {
    return new Promise((resolve, reject) => {
      const img = new Image();
      img.crossOrigin = 'Anonymous';
      img.onload = () => {
        const canvas = document.createElement('canvas');
        canvas.width = img.width;
        canvas.height = img.height;
        const ctx = canvas.getContext('2d');
        ctx.drawImage(img, 0, 0);
        const dataURL = canvas.toDataURL('image/png');
        resolve(dataURL);
      };
      img.onerror = reject;
      img.src = url;
    });
  }

  // Fallback: Generate HTML document
  generateHtmlDocument(videos, screenshots) {
    // ... existing HTML generation logic ...
  }
}
