
export const MIME_TYPES: { [key: string]: string } = {
    'png': 'image/png',
    'jpg': 'image/jpeg',
    'jpeg': 'image/jpeg',
    'gif': 'image/gif',
    'bmp': 'image/bmp',
    'tiff': 'image/tiff',
    'emf': 'image/emf',
    'wmf': 'image/wmf',
    'svg': 'image/svg+xml',
    'ico': 'image/x-icon',
    'tif': 'image/tiff',
    'wdp': 'image/vnd.ms-photo'
};

export const SUPPORTED_IMAGE_EXTENSIONS = Object.keys(MIME_TYPES);

// Extensions that are decoded as UTF-8 text; everything else is treated as binary
// to prevent corruption of EMF, WMF, OLE, and other binary assets in .vsdx archives.
const TEXT_EXTENSIONS = new Set(['xml', 'rels']);

export function isBinaryExtension(ext: string): boolean {
    return !TEXT_EXTENSIONS.has(ext.toLowerCase());
}
