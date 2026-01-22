
export const MIME_TYPES: { [key: string]: string } = {
    'png': 'image/png',
    'jpg': 'image/jpeg',
    'jpeg': 'image/jpeg',
    'gif': 'image/gif',
    'bmp': 'image/bmp',
    'tiff': 'image/tiff'
};

export const SUPPORTED_IMAGE_EXTENSIONS = Object.keys(MIME_TYPES);

export function isBinaryExtension(ext: string): boolean {
    return SUPPORTED_IMAGE_EXTENSIONS.includes(ext.toLowerCase());
}
