
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
    MASTER:  'http://schemas.microsoft.com/visio/2010/relationships/master',
    PAGES: 'http://schemas.microsoft.com/visio/2010/relationships/pages',
    PAGE: 'http://schemas.microsoft.com/visio/2010/relationships/page',
    WINDOWS: 'http://schemas.microsoft.com/visio/2010/relationships/windows',
    DOCUMENT: 'http://schemas.microsoft.com/visio/2010/relationships/document',
    CORE_PROPERTIES: 'http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties',
    EXTENDED_PROPERTIES: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties'
} as const;

/** Visio shape `Type` attribute values. */
export const SHAPE_TYPES = {
    Shape:   'Shape',
    Group:   'Group',
    Foreign: 'Foreign',
} as const;

/**
 * Visio ShapeSheet section names (the `N` attribute on `<Section>` elements).
 * Used when finding or filtering sections by name in page/master XML.
 */
export const SECTION_NAMES = {
    Line:       'Line',
    Fill:       'Fill',
    Character:  'Character',
    Paragraph:  'Paragraph',
    TextBlock:  'TextBlock',
    Geometry:   'Geometry',
    Connection: 'Connection',
    Property:   'Property',
    User:       'User',
    LayerMem:   'LayerMem',
    Layer:      'Layer',
} as const;

/**
 * Structural relationship types stored in `<Relationship>` elements
 * inside Visio page XML (distinct from OPC `.rels` relationship types).
 */
export const STRUCT_RELATIONSHIP_TYPES = {
    Container: 'Container',
} as const;

/**
 * Maps the user-facing `LengthUnit` strings to the Visio XML `Unit` attribute
 * values used in PageSheet drawing-scale cells (`PageScale`, `DrawingScale`).
 */
export const LENGTH_UNIT_TO_VISIO: Record<string, string> = {
    in:  'IN',
    ft:  'FT',
    yd:  'YD',
    mi:  'MI',
    mm:  'MM',
    cm:  'CM',
    m:   'M',
    km:  'KM',
} as const;

/** Reverse map: Visio unit string → `LengthUnit`. */
export const VISIO_TO_LENGTH_UNIT: Record<string, string> = {
    IN:  'in',
    FT:  'ft',
    YD:  'yd',
    MI:  'mi',
    MM:  'mm',
    CM:  'cm',
    M:   'm',
    KM:  'km',
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
    VISIO_MASTERS: 'application/vnd.ms-visio.masters+xml',
    VISIO_MASTER:  'application/vnd.ms-visio.master+xml',
    VISIO_PAGES: 'application/vnd.ms-visio.pages+xml',
    VISIO_PAGE: 'application/vnd.ms-visio.page+xml',
    VISIO_WINDOWS: 'application/vnd.ms-visio.windows+xml',
    CORE_PROPERTIES: 'application/vnd.openxmlformats-package.core-properties+xml',
    EXTENDED_PROPERTIES: 'application/vnd.openxmlformats-officedocument.extended-properties+xml'
} as const;
