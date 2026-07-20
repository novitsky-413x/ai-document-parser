import { Document, Packer, Paragraph, TextRun } from 'docx';
import { ensureParentDir, writeFileAtomic } from '../utils/paths.js';
import { wrapError } from '../utils/errors.js';

export async function writeDocx(
  filePath: string,
  content: string | string[]
): Promise<{ message: string }> {
  try {
    const paragraphs = Array.isArray(content)
      ? content
      : content.split(/\r?\n/);

    const children = paragraphs.map(
      (text) =>
        new Paragraph({
          children: [new TextRun(text || ' ')],
        })
    );

    const doc = new Document({
      sections: [
        {
          children:
            children.length > 0 ? children : [new Paragraph({ children: [] })],
        },
      ],
    });

    const buffer = await Packer.toBuffer(doc);
    await ensureParentDir(filePath);
    await writeFileAtomic(filePath, buffer);
    return {
      message: `Created DOCX: ${filePath} (${paragraphs.length} paragraph(s))`,
    };
  } catch (err) {
    throw wrapError('WRITE_FAILED', `Failed to write docx ${filePath}`, err);
  }
}
