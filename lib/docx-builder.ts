import PizZip from 'pizzip';
import sharp from 'sharp';

interface ImageData {
  name: string;
  buffer: Buffer;
  width: number;
  height: number;
  mediaName: string;
  rId: string;
  widthEmu: number;
  heightEmu: number;
}

type MaybeImage = ImageData | null;

// 1 dxa(twip) = 914400/1440 = 635 EMU
// 1px at 96DPI = 914400/96 = 9525 EMU
// → cell content width (px) = (colW_dxa - padding_dxa) * 635 / 9525
//                            = (colW_dxa - padding_dxa) / 15
const CELL_PADDING_DXA = 57 * 2; // 57 dxa each side
const TABLE_WIDTH_DXA = 9360;    // standard content width (6.5 inch)
const EMU_PER_PX = 9525;

export async function buildReport(
  docxBuffer: Buffer,
  photos: { name: string; buffer: Buffer }[],
  columns: number = 3
): Promise<Buffer> {
  const cols = Math.max(1, Math.min(6, columns)); // clamp 1–6

  const zip = new PizZip(docxBuffer);

  const docFile = zip.file('word/document.xml');
  const relsFile = zip.file('word/_rels/document.xml.rels');
  const contentTypesFile = zip.file('[Content_Types].xml');

  if (!docFile || !relsFile || !contentTypesFile) {
    throw new Error('올바른 .docx 파일이 아닙니다. (필수 XML 파일 누락)');
  }

  let docXml = docFile.asText();
  let relsXml = relsFile.asText();
  let contentTypesXml = contentTypesFile.asText();

  // Find the highest existing rId number to avoid collisions
  const existingIds = [...relsXml.matchAll(/\bId="rId(\d+)"/g)].map(m => parseInt(m[1], 10));
  let nextRId = existingIds.length > 0 ? Math.max(...existingIds) + 1 : 100;

  // Cell content width in pixels (used as target resize width)
  const colW = Math.floor(TABLE_WIDTH_DXA / cols);
  const contentWidthPx = Math.round((colW - CELL_PADDING_DXA) / 15);

  // Resize each image to fill cell width, maintaining aspect ratio
  const images: ImageData[] = await Promise.all(
    photos.map(async (photo, idx) => {
      const resized = await sharp(photo.buffer)
        .resize({ width: contentWidthPx, withoutEnlargement: false })
        .jpeg({ quality: 82 })
        .toBuffer();

      const meta = await sharp(resized).metadata();
      const w = meta.width ?? contentWidthPx;
      const h = meta.height ?? contentWidthPx;
      const rId = `rId${nextRId + idx}`;
      const mediaName = `photo_insert_${idx + 1}.jpeg`;

      return {
        name: photo.name,
        buffer: resized,
        width: w,
        height: h,
        mediaName,
        rId,
        widthEmu: w * EMU_PER_PX,
        heightEmu: h * EMU_PER_PX,
      };
    })
  );

  // Write resized images into word/media/ and register relationships
  for (const img of images) {
    zip.file(`word/media/${img.mediaName}`, img.buffer);

    const relEntry =
      `<Relationship Id="${img.rId}"` +
      ` Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"` +
      ` Target="media/${img.mediaName}"/>`;

    relsXml = relsXml.replace('</Relationships>', `  ${relEntry}\n</Relationships>`);
  }
  zip.file('word/_rels/document.xml.rels', relsXml);

  // Ensure JPEG content type is declared in [Content_Types].xml
  if (!contentTypesXml.includes('"jpeg"') && !contentTypesXml.includes("'jpeg'")) {
    contentTypesXml = contentTypesXml.replace(
      '</Types>',
      '  <Default Extension="jpeg" ContentType="image/jpeg"/>\n</Types>'
    );
    zip.file('[Content_Types].xml', contentTypesXml);
  }

  // Ensure drawing-related XML namespaces are declared on the document root
  docXml = ensureNamespaces(docXml);

  // Build the photo section and insert before the final <w:sectPr> (or before </w:body>)
  const sectionXml = buildPhotoSection(images, cols, colW);
  docXml = insertBeforeSectPr(docXml, sectionXml);

  zip.file('word/document.xml', docXml);

  return Buffer.from(
    zip.generate({
      type: 'nodebuffer',
      mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
      compression: 'DEFLATE',
    })
  );
}

/** Add required drawing namespaces to <w:document> if missing. */
function ensureNamespaces(docXml: string): string {
  const required: Array<[string, string]> = [
    ['xmlns:wp', 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'],
    ['xmlns:a', 'http://schemas.openxmlformats.org/drawingml/2006/main'],
    ['xmlns:pic', 'http://schemas.openxmlformats.org/drawingml/2006/picture'],
    ['xmlns:r', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'],
  ];

  for (const [attr, uri] of required) {
    if (!docXml.includes(`${attr}=`)) {
      docXml = docXml.replace(/(<w:document\b)/, `$1 ${attr}="${uri}"`);
    }
  }
  return docXml;
}

/**
 * Insert new XML content before the final <w:sectPr> (section properties)
 * which must remain the last child of <w:body>.
 */
function insertBeforeSectPr(docXml: string, insertXml: string): string {
  const lastSectPr = docXml.lastIndexOf('<w:sectPr');
  if (lastSectPr !== -1) {
    return docXml.slice(0, lastSectPr) + insertXml + '\n' + docXml.slice(lastSectPr);
  }
  return docXml.replace('</w:body>', insertXml + '\n</w:body>');
}

/** Build page-break + heading + N-column photo table XML. */
function buildPhotoSection(images: ImageData[], cols: number, colW: number): string {
  const rows: string[] = [];

  for (let i = 0; i < images.length; i += cols) {
    const chunk: MaybeImage[] = images.slice(i, i + cols);
    while (chunk.length < cols) chunk.push(null);

    rows.push(`<w:tr>${chunk.map((img, j) => buildImageCell(img, i + j + 1, colW)).join('')}</w:tr>`);
    rows.push(`<w:tr>${chunk.map(img => buildNameCell(img?.name ?? '', colW)).join('')}</w:tr>`);
  }

  const gridCols = Array.from({ length: cols }, () => `<w:gridCol w:w="${colW}"/>`).join('');

  const table = `<w:tbl>
  <w:tblPr>
    <w:tblW w:w="${TABLE_WIDTH_DXA}" w:type="dxa"/>
    <w:tblLayout w:type="fixed"/>
    <w:tblBorders>
      <w:top     w:val="single" w:sz="4" w:space="0" w:color="auto"/>
      <w:left    w:val="single" w:sz="4" w:space="0" w:color="auto"/>
      <w:bottom  w:val="single" w:sz="4" w:space="0" w:color="auto"/>
      <w:right   w:val="single" w:sz="4" w:space="0" w:color="auto"/>
      <w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/>
      <w:insideV w:val="single" w:sz="4" w:space="0" w:color="auto"/>
    </w:tblBorders>
    <w:tblCellMar>
      <w:top    w:w="57" w:type="dxa"/>
      <w:left   w:w="57" w:type="dxa"/>
      <w:bottom w:w="57" w:type="dxa"/>
      <w:right  w:w="57" w:type="dxa"/>
    </w:tblCellMar>
  </w:tblPr>
  <w:tblGrid>${gridCols}</w:tblGrid>
  ${rows.join('\n  ')}
</w:tbl>`;

  return (
    `<w:p><w:r><w:br w:type="page"/></w:r></w:p>\n` +
    `<w:p>
  <w:pPr>
    <w:jc w:val="center"/>
    <w:spacing w:before="240" w:after="240"/>
  </w:pPr>
  <w:r>
    <w:rPr><w:b/><w:sz w:val="36"/><w:szCs w:val="36"/></w:rPr>
    <w:t>사진 첨부</w:t>
  </w:r>
</w:p>\n` +
    table
  );
}

/** Cell containing a centered image that fills the cell width. */
function buildImageCell(img: MaybeImage, docPrId: number, colW: number): string {
  const content = img
    ? `<w:p>
    <w:pPr>
      <w:jc w:val="center"/>
      <w:spacing w:before="0" w:after="0"/>
    </w:pPr>
    <w:r>
      <w:drawing>
        <wp:inline distT="0" distB="0" distL="0" distR="0"
          xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">
          <wp:extent cx="${img.widthEmu}" cy="${img.heightEmu}"/>
          <wp:effectExtent l="0" t="0" r="0" b="0"/>
          <wp:docPr id="${docPrId}" name="Photo${docPrId}"/>
          <wp:cNvGraphicFramePr>
            <a:graphicFrameLocks
              xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
              noChangeAspect="1"/>
          </wp:cNvGraphicFramePr>
          <a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
            <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
              <pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
                <pic:nvPicPr>
                  <pic:cNvPr id="${docPrId}" name="Photo${docPrId}"/>
                  <pic:cNvPicPr>
                    <a:picLocks noChangeAspect="1" noChangeArrowheads="1"/>
                  </pic:cNvPicPr>
                </pic:nvPicPr>
                <pic:blipFill>
                  <a:blip r:embed="${img.rId}"
                    xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/>
                  <a:stretch><a:fillRect/></a:stretch>
                </pic:blipFill>
                <pic:spPr bwMode="auto">
                  <a:xfrm>
                    <a:off x="0" y="0"/>
                    <a:ext cx="${img.widthEmu}" cy="${img.heightEmu}"/>
                  </a:xfrm>
                  <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
                </pic:spPr>
              </pic:pic>
            </a:graphicData>
          </a:graphic>
        </wp:inline>
      </w:drawing>
    </w:r>
  </w:p>`
    : `<w:p><w:pPr><w:jc w:val="center"/></w:pPr></w:p>`;

  return `<w:tc>
  <w:tcPr><w:tcW w:w="${colW}" w:type="dxa"/></w:tcPr>
  ${content}
</w:tc>`;
}

/** Cell containing the original filename, centered. */
function buildNameCell(name: string, colW: number): string {
  return `<w:tc>
  <w:tcPr><w:tcW w:w="${colW}" w:type="dxa"/></w:tcPr>
  <w:p>
    <w:pPr>
      <w:jc w:val="center"/>
      <w:spacing w:before="0" w:after="60"/>
    </w:pPr>
    <w:r>
      <w:rPr><w:sz w:val="16"/><w:szCs w:val="16"/></w:rPr>
      <w:t xml:space="preserve">${escapeXml(name)}</w:t>
    </w:r>
  </w:p>
</w:tc>`;
}

function escapeXml(str: string): string {
  return str
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}
