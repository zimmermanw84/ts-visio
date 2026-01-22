
export const XML_NAMESPACES = {
    VISIO_MAIN: 'http://schemas.microsoft.com/office/visio/2012/main',
    RELATIONSHIPS: 'http://schemas.openxmlformats.org/package/2006/relationships',
    RELATIONSHIPS_OFFICE: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    CONTENT_TYPES: 'http://schemas.openxmlformats.org/package/2006/content-types',
    CORE_PROPERTIES: 'http://schemas.openxmlformats.org/package/2006/metadata/core-properties',
    EXTENDED_PROPERTIES: 'http://schemas.openxmlformats.org/officeDocument/2006/extended-properties',
    DOC_PROPS_VTYPES: 'http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes',
    DC_ELEMENTS: 'http://purl.org/dc/elements/1.1/',
    DC_TERMS: 'http://purl.org/dc/terms/',
    DC_DCMITYPE: 'http://purl.org/dc/dcmitype/',
    XSI: 'http://www.w3.org/2001/XMLSchema-instance'
} as const;

export const RELATIONSHIP_TYPES = {
    IMAGE: 'http://schemas.microsoft.com/office/2006/relationships/image',
    MASTERS: 'http://schemas.microsoft.com/visio/2010/relationships/masters',
    PAGES: 'http://schemas.microsoft.com/visio/2010/relationships/pages',
    PAGE: 'http://schemas.microsoft.com/visio/2010/relationships/page',
    WINDOWS: 'http://schemas.microsoft.com/visio/2010/relationships/windows',
    DOCUMENT: 'http://schemas.microsoft.com/visio/2010/relationships/document',
    CORE_PROPERTIES: 'http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties',
    EXTENDED_PROPERTIES: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties'
} as const;

export const CONTENT_TYPES = {
    PNG: 'image/png',
    JPEG: 'image/jpeg',
    GIF: 'image/gif',
    BMP: 'image/bmp',
    TIFF: 'image/tiff',
    XML_RELATIONSHIPS: 'application/vnd.openxmlformats-package.relationships+xml',
    XML: 'application/xml',
    VISIO_DRAWING: 'application/vnd.ms-visio.drawing.main+xml',
    VISIO_PAGES: 'application/vnd.ms-visio.pages+xml',
    VISIO_PAGE: 'application/vnd.ms-visio.page+xml',
    VISIO_WINDOWS: 'application/vnd.ms-visio.windows+xml',
    CORE_PROPERTIES: 'application/vnd.openxmlformats-package.core-properties+xml',
    EXTENDED_PROPERTIES: 'application/vnd.openxmlformats-officedocument.extended-properties+xml'
} as const;
