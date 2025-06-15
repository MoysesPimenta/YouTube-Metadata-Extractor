// File: js/document-generator.js

// Document Generator with DOCX Support and YouTube Page Screenshots
tag://javascript
class DocumentGenerator {
  constructor(screenshotApiToken) {
    this.xlsx = null;
    this.docx = null;
    this.fileSaver = null;
    // Token for third-party screenshot service (e.g. ScreenshotAPI.net)
    this.screenshotApiToken = screenshotApiToken;
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
        await this.loadScript('https://cdn.jsdelivr.net/npm/docx@8.5.0/build/index.umd.js');
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

      return { blob, filename: 'playlist_data.xlsx' };
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
    return { blob, filename: 'playlist_data.csv' };
  }

  // Generate Word document (DOCX format) with actual YouTube page screenshots
  async generateWordDocument(videos) {
    if (!videos || videos.length === 0) {
      throw new Error('Nenhum dado disponível para gerar o documento.');
    }

    try {
      if (!this.docx) {
        console.warn('DOCX library not loaded, falling back to HTML');
        return this.generateHtmlDocument(videos);
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
        HeadingLevel,
        PageBreak,
        Packer,
      } = this.docx;

      const doc = new Document({
        creator: 'YouTube Playlist Extractor',
        title: 'Comprovação de Dados - Playlist YouTube',
        description: 'Dados extraídos de playlist do YouTube com screenshots das páginas',
      });

      // Title section
      doc.addSection({
        properties: {},
        children: [
          new Paragraph({
            text: 'Comprovação de Dados - Playlist YouTube',
            heading: HeadingLevel.HEADING_1,
            alignment: AlignmentType.CENTER,
          }),
        ],
      });

      for (let i = 0; i < videos.length; i++) {
        const video = videos[i];
        const videoUrl = `https://www.youtube.com/watch?v=${video.videoId}`;

        // Fetch actual page screenshot
        let screenshotData;
        try {
          screenshotData = await this.fetchPageScreenshot(videoUrl);
        } catch (e) {
          console.error('Failed to fetch page screenshot:', e);
        }

        const sectionChildren = [];
        sectionChildren.push(
          new Paragraph({
            text: `Vídeo ${i + 1}: ${video.title}`,
            heading: HeadingLevel.HEADING_2,
          })
        );

        // Metadata table
        sectionChildren.push(
          new Table({
            rows: [
              this._makeRow('Duração', video.duration),
              this._makeRow('Views', video.views.toString()),
              this._makeRow('Likes', video.likes.toString()),
              this._makeRow('Link', videoUrl),
            ],
            width: { size: 100, type: 'pct' },
            borders: {
              top: { style: BorderStyle.SINGLE, size: 1, color: 'CCCCCC' },
              bottom: { style: BorderStyle.SINGLE, size: 1, color: 'CCCCCC' },
              left: { style: BorderStyle.SINGLE, size: 1, color: 'CCCCCC' },
              right: { style: BorderStyle.SINGLE, size: 1, color: 'CCCCCC' },
              insideHorizontal: { style: BorderStyle.NONE },
              insideVertical: { style: BorderStyle.NONE },
            },
          })
        );

        // Add screenshot if available
        if (screenshotData) {
          sectionChildren.push(
            new Paragraph({
              children: [
                new ImageRun({
                  data: screenshotData.split(',')[1],
                  transformation: { width: 600, height: 338 },
                }),
              ],
            })
          );
        }

        // Page break unless last
        if (i < videos.length - 1) {
          sectionChildren.push(new Paragraph({ children: [new PageBreak()] }));
        }

        doc.addSection({ properties: {}, children: sectionChildren });
      }

      const buffer = await Packer.toBuffer(doc);
      const blob = new Blob([buffer], {
        type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
      });
      return { blob, filename: 'comprovacao_videos.docx' };
    } catch (error) {
      console.error('Error generating DOCX:', error);
      return this.generateHtmlDocument(videos);
    }
  }

  // Fetch screenshot of the YouTube video page via ScreenshotAPI.net (or similar)
  async fetchPageScreenshot(url) {
    if (!this.screenshotApiToken) {
      throw new Error('Screenshot API token is missing.');
    }
    const apiUrl = `https://api.screenshotapi.net/screenshot?token=${this.screenshotApiToken}&url=${encodeURIComponent(
      url
    )}&full_page=true`; // full page capture

    const response = await fetch(apiUrl);
    if (!response.ok) {
      throw new Error(`Screenshot API responded with ${response.status}`);
    }
    const result = await response.json();
    // result.screenshot contains the Base64-encoded image data
    return `data:image/png;base64,${result.screenshot}`;
  }

  // Helper to create a metadata table row
  _makeRow(label, value) {
    const { TableRow, TableCell, Paragraph, BorderStyle } = this.docx;
    return new TableRow({
      children: [
        new TableCell({
          width: { size: 30, type: 'pct' },
          children: [new Paragraph({ text: `${label}:`, bold: true })],
        }),
        new TableCell({
          width: { size: 70, type: 'pct' },
          children: [new Paragraph({ text: value })],
        }),
      ],
    });
  }

  // Fallback: Generate HTML document
  generateHtmlDocument(videos) {
    let html = `<!DOCTYPE html><html><head><meta charset="UTF-8"><title>Comprovação de Dados</title></head><body>`;
    videos.forEach((video, i) => {
      html += `<h2>Vídeo ${i + 1}: ${video.title}</h2>`;
      html += `<ul>`;
      html += `<li><strong>Duração:</strong> ${video.duration}</li>`;
      html += `<li><strong>Views:</strong> ${video.views}</li>`;
      html += `<li><strong>Likes:</strong> ${video.likes}</li>`;
      html += `<li><strong>Link:</strong> <a href="https://www.youtube.com/watch?v=${video.videoId}">Ver no YouTube</a></li>`;
      html += `</ul>`;
      html += `<img src="https://img.youtube.com/vi/${video.videoId}/maxresdefault.jpg" style="max-width:600px;"><hr>`;
    });
    html += `</body></html>`;

    const blob = new Blob([html], { type: 'text/html;charset=utf-8;' });
    return { blob, filename: 'comprovacao_videos.html' };
  }
}
