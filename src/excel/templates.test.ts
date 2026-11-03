import * as fs from 'fs';
import * as path from 'path';
// eslint-disable-next-line @typescript-eslint/no-require-imports
const JSZip = require('jszip');

const root = path.resolve(__dirname, '../..');
const logo = fs.readFileSync(path.join(root, 'src/webparts/timesheet/assets/peak-logo.jpg'));

// JPEG size from the first SOFn marker.
function jpegSize(b: Buffer): { w: number; h: number } {
  let i = 2;
  while (i < b.length) {
    const marker = b[i + 1];
    if (marker >= 0xc0 && marker <= 0xcf && marker !== 0xc4 && marker !== 0xc8 && marker !== 0xcc) {
      return { h: b.readUInt16BE(i + 5), w: b.readUInt16BE(i + 7) };
    }
    i += 2 + b.readUInt16BE(i + 2);
  }
  throw new Error('no SOF marker');
}

describe.each(['Admin', 'Instructor', 'Medic'])('%s template logo', (type) => {
  let zip: { files: Record<string, unknown>; file(n: string): { async(t: string): Promise<never> } };
  beforeAll(async () => {
    zip = await JSZip.loadAsync(fs.readFileSync(path.join(root, 'Templates/xlsx', `${type} Timesheet Template.xlsx`)));
  });

  it('has only the PEAK-only logo', async () => {
    const media = Object.keys(zip.files).filter((n) => n.startsWith('xl/media/'));
    expect(media).toHaveLength(1);
    const bytes: Buffer = await zip.file(media[0]).async('nodebuffer');
    expect(bytes.equals(logo)).toBe(true);
  });

  it('draws it uncropped at the logo aspect ratio', async () => {
    const xml: string = await zip.file('xl/drawings/drawing1.xml').async('string');
    expect(xml).not.toContain('srcRect');
    const m = /<a:ext cx="(\d+)" cy="(\d+)"\/>/.exec(xml);
    const { w, h } = jpegSize(logo);
    expect(Number(m![1]) / Number(m![2])).toBeCloseTo(w / h, 1);
    expect(Math.abs(Number(m![1]) / Number(m![2]) / (w / h) - 1)).toBeLessThan(0.02);
  });
});
