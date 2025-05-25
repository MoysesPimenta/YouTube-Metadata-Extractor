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
        await this.loadScript('/js/docx.min.js');
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
      // If SheetJS is loaded, use it
      if (this.xlsx) {
        // Prepare data for Excel
        const data = videos.map(video => ({
          'Nome do Episodio': video.title,
          'Duração': video.duration,
          'Views': video.views,
          'Likes': video.likes,
          'Link': `https://www.youtube.com/watch?v=${video.videoId}`,
          'Data de Publicacao': new Date(video.publishedDate).toLocaleDateString()
        }));

        // Create workbook and worksheet
        const ws = this.xlsx.utils.json_to_sheet(data);
        const wb = this.xlsx.utils.book_new();
        this.xlsx.utils.book_append_sheet(wb, ws, 'Playlist');

        // Generate Excel file
        const excelBuffer = this.xlsx.write(wb, { bookType: 'xlsx', type: 'array' });
        const blob = new Blob([excelBuffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
        
        return {
          blob,
          filename: 'playlist_data.xlsx'
        };
      } else {
        // Fallback to CSV if SheetJS isn't loaded
        return this.generateCSV(videos);
      }
    } catch (error) {
      console.error('Error generating Excel:', error);
      // Fallback to CSV
      return this.generateCSV(videos);
    }
  }

  // Fallback: Generate CSV file
  generateCSV(videos) {
    // Create CSV header
    let csv = 'Nome do Episodio,Duração,Views,Likes,Link,Data de Publicacao\n';
    
    // Add data rows
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
    
    // Create blob
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
      // Check if docx library is loaded
      if (!this.docx) {
        console.warn('DOCX library not loaded, falling back to HTML');
        return this.generateHtmlDocument(videos, screenshots);
      }
      
      const { Document, Paragraph, TextRun, ImageRun, Table, TableRow, TableCell, BorderStyle, AlignmentType, HeadingLevel, PageBreak } = this.docx;
      
      // Create document
      const doc = new Document({
        creator: "YouTube Playlist Extractor",
        title: "Comprovação de Dados - Playlist YouTube",
        description: "Documento de comprovação de dados extraídos de uma playlist do YouTube",
        styles: {
          paragraphStyles: [
            {
              id: "Heading1",
              name: "Heading 1",
              basedOn: "Normal",
              next: "Normal",
              quickFormat: true,
              run: {
                size: 36,
                bold: true,
                color: "2B579A"
              },
              paragraph: {
                spacing: {
                  after: 240
                }
              }
            },
            {
              id: "Heading2",
              name: "Heading 2",
              basedOn: "Normal",
              next: "Normal",
              quickFormat: true,
              run: {
                size: 28,
                bold: true,
                color: "2B579A"
              },
              paragraph: {
                spacing: {
                  before: 240,
                  after: 120
                }
              }
            },
            {
              id: "VideoTitle",
              name: "Video Title",
              basedOn: "Normal",
              next: "Normal",
              quickFormat: true,
              run: {
                size: 24,
                bold: true,
                color: "202124"
              },
              paragraph: {
                spacing: {
                  before: 240,
                  after: 120
                }
              }
            },
            {
              id: "MetadataLabel",
              name: "Metadata Label",
              basedOn: "Normal",
              next: "Normal",
              quickFormat: true,
              run: {
                size: 22,
                bold: true
              }
            },
            {
              id: "MetadataValue",
              name: "Metadata Value",
              basedOn: "Normal",
              next: "Normal",
              quickFormat: true,
              run: {
                size: 22,
                color: "202124"
              },
              paragraph: {
                spacing: {
                  after: 120
                }
              }
            },
            {
              id: "HighlightedValue",
              name: "Highlighted Value",
              basedOn: "Normal",
              next: "Normal",
              quickFormat: true,
              run: {
                size: 22,
                bold: true,
                color: "C00000"  // Red color for highlighted values
              }
            }
          ]
        }
      });
      
      // Add title
      doc.addSection({
        properties: {},
        children: [
          new Paragraph({
            text: "Comprovação de Dados - Playlist YouTube",
            heading: HeadingLevel.HEADING_1,
            alignment: AlignmentType.CENTER
          })
        ]
      });
      
      // Process each video
      for (let i = 0; i < videos.length; i++) {
        const video = videos[i];
        const screenshot = screenshots[i];
        
        // Convert screenshot URL to image data
        let imageData;
        try {
          imageData = await this.getImageDataFromUrl(screenshot.imageUrl);
        } catch (error) {
          console.error('Error loading image:', error);
          imageData = null;
        }
        
        // Add video section
        const videoSection = [];
        
        // Add video title
        videoSection.push(
          new Paragraph({
            text: `Vídeo ${i + 1}: ${video.title}`,
            heading: HeadingLevel.HEADING_2
          })
        );
        
        // Add metadata table
        const metadataTable = new Table({
          width: {
            size: 100,
            type: "pct"
          },
          borders: {
            top: { style: BorderStyle.NONE },
            bottom: { style: BorderStyle.NONE },
            left: { style: BorderStyle.NONE },
            right: { style: BorderStyle.NONE },
            insideHorizontal: { style: BorderStyle.NONE },
            insideVertical: { style: BorderStyle.NONE }
          },
          rows: [
            // Nome do Episódio
            new TableRow({
              children: [
                new TableCell({
                  width: {
                    size: 30,
                    type: "pct"
                  },
                  children: [
                    new Paragraph({
                      text: "Nome do Episódio:",
                      style: "MetadataLabel"
                    })
                  ]
                }),
                new TableCell({
                  width: {
                    size: 70,
                    type: "pct"
                  },
                  children: [
                    new Paragraph({
                      text: video.title,
                      style: "MetadataValue"
                    })
                  ]
                })
              ]
            }),
            // Duração
            new TableRow({
              children: [
                new TableCell({
                  width: {
                    size: 30,
                    type: "pct"
                  },
                  children: [
                    new Paragraph({
                      text: "Duração:",
                      style: "MetadataLabel"
                    })
                  ]
                }),
                new TableCell({
                  width: {
                    size: 70,
                    type: "pct"
                  },
                  children: [
                    new Paragraph({
                      text: video.duration,
                      style: "MetadataValue"
                    })
                  ]
                })
              ]
            }),
            // Views (Highlighted)
            new TableRow({
              children: [
                new TableCell({
                  width: {
                    size: 30,
                    type: "pct"
                  },
                  children: [
                    new Paragraph({
                      text: "Views:",
                      style: "MetadataLabel"
                    })
                  ]
                }),
                new TableCell({
                  width: {
                    size: 70,
                    type: "pct"
                  },
                  children: [
                    new Paragraph({
                      children: [
                        new TextRun({
                          text: video.views.toLocaleString(),
                          style: "HighlightedValue"
                        })
                      ]
                    })
                  ]
                })
              ]
            }),
            // Likes (Highlighted)
            new TableRow({
              children: [
                new TableCell({
                  width: {
                    size: 30,
                    type: "pct"
                  },
                  children: [
                    new Paragraph({
                      text: "Likes:",
                      style: "MetadataLabel"
                    })
                  ]
                }),
                new TableCell({
                  width: {
                    size: 70,
                    type: "pct"
                  },
                  children: [
                    new Paragraph({
                      children: [
                        new TextRun({
                          text: video.likes.toLocaleString(),
                          style: "HighlightedValue"
                        })
                      ]
                    })
                  ]
                })
              ]
            }),
            // Link (Highlighted)
            new TableRow({
              children: [
                new TableCell({
                  width: {
                    size: 30,
                    type: "pct"
                  },
                  children: [
                    new Paragraph({
                      text: "Link:",
                      style: "MetadataLabel"
                    })
                  ]
                }),
                new TableCell({
                  width: {
                    size: 70,
                    type: "pct"
                  },
                  children: [
                    new Paragraph({
                      children: [
                        new TextRun({
                          text: `https://www.youtube.com/watch?v=${video.videoId}`,
                          style: "HighlightedValue"
                        })
                      ]
                    })
                  ]
                })
              ]
            }),
            // Data de Publicação (Highlighted)
            new TableRow({
              children: [
                new TableCell({
                  width: {
                    size: 30,
                    type: "pct"
                  },
                  children: [
                    new Paragraph({
                      text: "Data de Publicação:",
                      style: "MetadataLabel"
                    })
                  ]
                }),
                new TableCell({
                  width: {
                    size: 70,
                    type: "pct"
                  },
                  children: [
                    new Paragraph({
                      children: [
                        new TextRun({
                          text: new Date(video.publishedDate).toLocaleDateString(),
                          style: "HighlightedValue"
                        })
                      ]
                    })
                  ]
                })
              ]
            })
          ]
        });
        
        videoSection.push(metadataTable);
        
        // Add screenshot
        if (imageData) {
          videoSection.push(
            new Paragraph({
              text: "Captura de tela (comprovação visual):",
              style: "MetadataLabel",
              spacing: {
                before: 240,
                after: 120
              }
            })
          );
          
          videoSection.push(
            new Paragraph({
              children: [
                new ImageRun({
                  data: imageData,
                  transformation: {
                    width: 600,
                    height: 350
                  }
                })
              ],
              spacing: {
                after: 240
              }
            })
          );
        }
        
        // Add page break if not the last video
        if (i < videos.length - 1) {
          videoSection.push(new Paragraph({ children: [new PageBreak()] }));
        }
        
        // Add video section to document
        doc.addSection({
          properties: {},
          children: videoSection
        });
      }
      
      // Generate DOCX file
      const buffer = await Packer.toBuffer(doc);
      const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
      
      return {
        blob,
        filename: 'comprovacao_videos.docx'
      };
    } catch (error) {
      console.error('Error generating DOCX:', error);
      // Fallback to HTML if DOCX generation fails
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
        
        // Convert to binary data
        const dataUrl = canvas.toDataURL('image/png');
        const base64 = dataUrl.split(',')[1];
        resolve(base64);
      };
      
      img.onerror = () => {
        reject(new Error('Failed to load image'));
      };
      
      img.src = url;
    });
  }

  // Fallback: Generate HTML document
  generateHtmlDocument(videos, screenshots) {
    if (!videos || videos.length === 0 || !screenshots || screenshots.length === 0) {
      throw new Error('Nenhum dado disponível para gerar o documento.');
    }
    
    // Create HTML document
    let html = `
      <!DOCTYPE html>
      <html>
      <head>
        <meta charset="UTF-8">
        <title>Comprovação de Dados - Playlist YouTube</title>
        <style>
          body { font-family: Arial, sans-serif; margin: 40px; }
          h1 { color: #2B579A; text-align: center; font-size: 24pt; margin-bottom: 30px; }
          .video { margin-bottom: 30px; border-bottom: 1px solid #eee; padding-bottom: 20px; }
          .video h2 { color: #2B579A; font-size: 18pt; }
          .metadata { margin-bottom: 15px; }
          .metadata p { margin: 5px 0; font-size: 12pt; }
          .metadata strong { font-weight: bold; }
          .highlighted { color: #C00000; font-weight: bold; }
          img { max-width: 100%; border: 1px solid #ddd; }
          .page-break { page-break-after: always; }
        </style>
      </head>
      <body>
        <h1>Comprovação de Dados - Playlist YouTube</h1>
    `;
    
    // Add each video
    videos.forEach((video, index) => {
      const screenshot = screenshots[index] || { imageUrl: '' };
      
      html += `
        <div class="video">
          <h2>Vídeo ${index + 1}: ${video.title}</h2>
          <div class="metadata">
            <p><strong>Nome do Episódio:</strong> ${video.title}</p>
            <p><strong>Duração:</strong> ${video.duration}</p>
            <p><strong>Views:</strong> <span class="highlighted">${video.views.toLocaleString()}</span></p>
            <p><strong>Likes:</strong> <span class="highlighted">${video.likes.toLocaleString()}</span></p>
            <p><strong>Link:</strong> <span class="highlighted">https://www.youtube.com/watch?v=${video.videoId}</span></p>
            <p><strong>Data de Publicação:</strong> <span class="highlighted">${new Date(video.publishedDate).toLocaleDateString()}</span></p>
          </div>
          <p><strong>Captura de tela (comprovação visual):</strong></p>
          <img src="${screenshot.imageUrl}" alt="Captura de tela: ${video.title}">
        </div>
        ${index < videos.length - 1 ? '<div class="page-break"></div>' : ''}
      `;
    });
    
    html += `
      </body>
      </html>
    `;
    
    // Create blob
    const blob = new Blob([html], { type: 'text/html;charset=utf-8;' });
    
    return {
      blob,
      filename: 'comprovacao_videos.html'
    };
  }
}
