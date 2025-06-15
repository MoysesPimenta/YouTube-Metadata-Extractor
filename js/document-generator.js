// File: js/document-generator.js

// Document Generator with DOCX Support and Screenshot API
class DocumentGenerator {
  constructor(screenshotToken) {
    this.xlsx = null;
    this.docx = null;
    this.screenshotToken = screenshotToken;
    this.fileSaver = null;
    this.loadLibraries();
  }

  // Update the screenshot token at runtime
  setScreenshotToken(token) {
    this.screenshotToken = token;
  }

  // Load required libraries
  async loadLibraries() {
    // Load SheetJS for Excel
    if (!window.XLSX) {
      await this.loadScript('https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js');
    }
    this.xlsx = window.XLSX;

    // Load docx for Word
    if (!window.docx) {
      await this.loadScript('https://cdn.jsdelivr.net/npm/docx@8.5.0/build/index.umd.js');
    }
    this.docx = window.docx;

    // Load FileSaver for downloads
    if (!window.saveAs) {
      await this.loadScript('https://cdn.jsdelivr.net/npm/file-saver@2.0.5/dist/FileSaver.min.js');
    }
    this.fileSaver = window.saveAs;
  }

  // Helper to inject script tags
  loadScript(src) {
    return new Promise((resolve, reject) => {
      const s = document.createElement('script');
      s.src = src;
      s.onload = resolve;
      s.onerror = reject;
      document.head.appendChild(s);
    });
  }

  // Excel generation
  generateExcel(videos) {
    if (!videos || videos.length === 0) {
      throw new Error('Nenhum dado disponível para gerar a planilha.');
    }
    const data = videos.map(v => ({
      'Nome do Episodio': v.title,
      'Duração': v.duration,
      'Views': v.views,
      'Likes': v.likes,
      'Link': `https://www.youtube.com/watch?v=${v.videoId}`,
      'Data de Publicacao': new Date(v.publishedDate).toLocaleDateString()
    }));
    const ws = this.xlsx.utils.json_to_sheet(data);
    const wb = this.xlsx.utils.book_new();
    this.xlsx.utils.book_append_sheet(wb, ws, 'Playlist');
    const buffer = this.xlsx.write(wb, { bookType: 'xlsx', type: 'array' });
    const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    return { blob, filename: 'playlist_data.xlsx' };
  }

  // DOCX generation with proper addSection usage
  async generateWordDocument(videos) {
    // Ensure library loaded
    if (!this.docx) {
      console.warn('DOCX library not loaded; falling back to HTML');
      return this.generateHtmlDocument(videos);
    }

    const {
      Document,
      Packer,
      Paragraph,
      Table,
      TableRow,
      TableCell,
      BorderStyle,
      AlignmentType,
      HeadingLevel,
      PageBreak,
      ImageRun
    } = this.docx;

    // Create document
    const doc = new Document({
      creator: 'YouTube Playlist Extractor',
      title: 'Comprovação de Dados - Playlist YouTube',
      description: 'Dados extraídos de playlist do YouTube'
    });

    // Title section
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

    // Add each video as a section
    for (let i = 0; i < videos.length; i++) {
      const video = videos[i];
      const url = `https://www.youtube.com/watch?v=${video.videoId}`;
      const children = [];

      // Video header
      children.push(
        new Paragraph({
          text: `Vídeo ${i + 1}: ${video.title}`,
          heading: HeadingLevel.HEADING_2
        })
      );

      // Metadata table
      children.push(
        new Table({
          rows: [
            this._row('Duração', video.duration),
            this._row('Views', String(video.views)),
            this._row('Likes', String(video.likes)),
            this._row('Link', url)
          ],
          width: { size: 100, type: 'pct' },
          borders: {
            top: { style: BorderStyle.SINGLE, size: 1, color: 'CCCCCC' },
            bottom: { style: BorderStyle.SINGLE, size: 1, color: 'CCCCCC' },
            left: { style: BorderStyle.SINGLE, size: 1, color: 'CCCCCC' },
            right: { style: BorderStyle.SINGLE, size: 1, color: 'CCCCCC' },
            insideHorizontal: { style: BorderStyle.NONE },
            insideVertical: { style: BorderStyle.NONE }
          }
        })
      );

      // Screenshot embed
      if (this.screenshotToken) {
        try {
          const imgData = await this._fetchScreenshot(url);
          children.push(
            new Paragraph({
              children: [
                new ImageRun({
                  data: imgData.split(',')[1],
                  transformation: { width: 600, height: 338 }
                })
              ]
            })
          );
        } catch (e) {
          console.warn('Screenshot fetch failed:', e);
        }
      }

      // Page break except after last
      if (i < videos.length - 1) {
        children.push(new Paragraph({ children: [ new PageBreak() ] }));
      }

      doc.addSection({ properties: {}, children });
    }

    // Pack to buffer and return
    const buffer = await Packer.toBuffer(doc);
    const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
    return { blob, filename: 'comprovacao_videos.docx' };
  }

  // Helper to create table rows
  _row(label, value) {
    const { TableRow, TableCell, Paragraph } = this.docx;
    return new TableRow({
      children: [
        new TableCell({
          width: { size: 30, type: 'pct' },
          children: [ new Paragraph({ text: `${label}:`, bold: true }) ]
        }),
        new TableCell({
          width: { size: 70, type: 'pct' },
          children: [ new Paragraph({ text: value }) ]
        })
      ]
    });
  }

  // Fetch full-page screenshot via API
  async _fetchScreenshot(url) {
    const apiUrl = `https://api.screenshotapi.net/screenshot?token=${this.screenshotToken}&url=${encodeURIComponent(url)}&full_page=true`;
    const res = await fetch(apiUrl);
    if (!res.ok) throw new Error(`Screenshot API error: ${res.status}`);
    const { screenshot } = await res.json();
    return `data:image/png;base64,${screenshot}`;
  }

  // Fallback HTML document
  generateHtmlDocument(videos) {
    let html = '<!DOCTYPE html><html><head><meta charset="UTF-8"><title>Comprovação de Dados</title></head><body>';
    videos.forEach((video, i) => {
      const thumb = `https://img.youtube.com/vi/${video.videoId}/hqdefault.jpg`;
      html += `<h2>Vídeo ${i+1}: ${video.title}</h2>`;
      html += `<ul><li><strong>Duração:</strong> ${video.duration}</li><li><strong>Views:</strong> ${video.views}</li><li><strong>Likes:</strong> ${video.likes}</li><li><strong>Link:</strong> <a href="https://www.youtube.com/watch?v=${video.videoId}">YouTube</a></li></ul>`;
      html += `<img src="${thumb}" style="max-width:600px;"><hr>`;
    });
    html += '</body></html>';
    const blob = new Blob([html], { type: 'text/html;charset=utf-8;' });
    return { blob, filename: 'comprovacao_videos.html' };
  }
}
