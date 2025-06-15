// File: js/document-generator.js

// Document Generator with correct docx sections usage and Screenshot API
document-generator.js
class DocumentGenerator {
  constructor(screenshotToken) {
    this.xlsx = null;
    this.docx = null;
    this.screenshotToken = screenshotToken;
    this.loadLibraries();
  }

  setScreenshotToken(token) {
    this.screenshotToken = token;
  }

  async loadLibraries() {
    if (!window.XLSX) await this.loadScript('https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js');
    this.xlsx = window.XLSX;
    if (!window.docx) await this.loadScript('https://cdn.jsdelivr.net/npm/docx@8.5.0/build/index.umd.js');
    this.docx = window.docx;
    if (!window.saveAs) await this.loadScript('https://cdn.jsdelivr.net/npm/file-saver@2.0.5/dist/FileSaver.min.js');
  }

  loadScript(src) {
    return new Promise((res, rej) => {
      const s = document.createElement('script'); s.src = src; s.onload = res; s.onerror = rej; document.head.appendChild(s);
    });
  }

  generateExcel(videos) {
    const data = videos.map(v => ({ 'Nome do Episodio': v.title, 'Duração': v.duration, 'Views': v.views, 'Likes': v.likes, 'Link': `https://www.youtube.com/watch?v=${v.videoId}`, 'Data de Publicacao': new Date(v.publishedDate).toLocaleDateString() }));
    const ws = this.xlsx.utils.json_to_sheet(data);
    const wb = this.xlsx.utils.book_new(); this.xlsx.utils.book_append_sheet(wb, ws, 'Playlist');
    const buf = this.xlsx.write(wb, { bookType: 'xlsx', type: 'array' });
    return { blob: new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }), filename: 'playlist_data.xlsx' };
  }

  async generateWordDocument(videos) {
    const { Document, Paragraph, Table, TableRow, TableCell, BorderStyle, AlignmentType, HeadingLevel, PageBreak, ImageRun, Packer } = this.docx;
    const sections = [];

    // Title section
    sections.push({
      properties: {},
      children: [
        new Paragraph({ text: 'Comprovação de Dados - Playlist YouTube', heading: HeadingLevel.HEADING_1, alignment: AlignmentType.CENTER })
      ]
    });

    // Build each video section
    for (let i = 0; i < videos.length; i++) {
      const v = videos[i];
      const url = `https://www.youtube.com/watch?v=${v.videoId}`;
      const children = [];

      // Header
      children.push(new Paragraph({ text: `Vídeo ${i + 1}: ${v.title}`, heading: HeadingLevel.HEADING_2 }));

      // Metadata table
      children.push(new Table({
        rows: [
          this._row('Duração', v.duration),
          this._row('Views', String(v.views)),
          this._row('Likes', String(v.likes)),
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
      }));

      // Screenshot image
      if (this.screenshotToken) {
        try {
          const imgData = await this._fetchScreenshot(url);
          children.push(new Paragraph({ children: [ new ImageRun({ data: imgData.split(',')[1], transformation: { width: 600, height: 338 } }) ] }));
        } catch (e) {
          console.warn('Screenshot error', e);
        }
      }

      // Page break if not last
      if (i < videos.length - 1) children.push(new Paragraph({ children: [ new PageBreak() ] }));

      sections.push({ properties: {}, children });
    }

    // Create document with sections
    const doc = new Document({ sections });
    const buffer = await Packer.toBuffer(doc);
    return { blob: new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' }), filename: 'comprovacao_videos.docx' };
  }

  _row(label, val) {
    const { TableRow, TableCell, Paragraph } = this.docx;
    return new TableRow({
      children: [
        new TableCell({ width: { size: 30, type: 'pct' }, children: [ new Paragraph({ text: `${label}:`, bold: true }) ] }),
        new TableCell({ width: { size: 70, type: 'pct' }, children: [ new Paragraph({ text: val }) ] })
      ]
    });
  }

  async _fetchScreenshot(url) {
    const api = `https://api.screenshotapi.net/screenshot?token=${this.screenshotToken}&url=${encodeURIComponent(url)}&full_page=true`;
    const res = await fetch(api);
    if (!res.ok) throw new Error(`Screenshot API error: ${res.status}`);
    const json = await res.json();
    return `data:image/png;base64,${json.screenshot}`;
  }

  generateHtmlDocument(videos) {
    let html = '<!DOCTYPE html><html><head><meta charset="UTF-8"><title>Comprovação de Dados</title></head><body>';
    videos.forEach((v, i) => {
      html += `<h2>Vídeo ${i+1}: ${v.title}</h2><ul><li><strong>Duração:</strong> ${v.duration}</li><li><strong>Views:</strong> ${v.views}</li><li><strong>Likes:</strong> ${v.likes}</li><li><strong>Link:</strong> <a href="https://www.youtube.com/watch?v=${v.videoId}">YouTube</a></li></ul><hr>`;
    });
    html += '</body></html>';
    return { blob: new Blob([html], { type: 'text/html;charset=utf-8;' }), filename: 'comprovacao_videos.html' };
  }
}
