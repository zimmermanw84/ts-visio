import { XML_NAMESPACES, RELATIONSHIP_TYPES, CONTENT_TYPES } from '../core/VisioConstants';

export const CONTENT_TYPES_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="${XML_NAMESPACES.CONTENT_TYPES}">
    <Default Extension="rels" ContentType="${CONTENT_TYPES.XML_RELATIONSHIPS}"/>
    <Default Extension="xml" ContentType="${CONTENT_TYPES.XML}"/>
    <Override PartName="/visio/document.xml" ContentType="${CONTENT_TYPES.VISIO_DRAWING}"/>
    <Override PartName="/visio/pages/pages.xml" ContentType="${CONTENT_TYPES.VISIO_PAGES}"/>
    <Override PartName="/visio/pages/page1.xml" ContentType="${CONTENT_TYPES.VISIO_PAGE}"/>
    <Override PartName="/visio/windows.xml" ContentType="${CONTENT_TYPES.VISIO_WINDOWS}"/>
    <Override PartName="/docProps/core.xml" ContentType="${CONTENT_TYPES.CORE_PROPERTIES}"/>
    <Override PartName="/docProps/app.xml" ContentType="${CONTENT_TYPES.EXTENDED_PROPERTIES}"/>
</Types>`;

export const RELS_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="${XML_NAMESPACES.RELATIONSHIPS}">
    <Relationship Id="rId1" Type="${RELATIONSHIP_TYPES.DOCUMENT}" Target="visio/document.xml"/>
    <Relationship Id="rId2" Type="${RELATIONSHIP_TYPES.CORE_PROPERTIES}" Target="docProps/core.xml"/>
    <Relationship Id="rId3" Type="${RELATIONSHIP_TYPES.EXTENDED_PROPERTIES}" Target="docProps/app.xml"/>
</Relationships>`;

export const DOCUMENT_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<VisioDocument xmlns="${XML_NAMESPACES.VISIO_MAIN}" xmlns:r="${XML_NAMESPACES.RELATIONSHIPS_OFFICE}" xml:space="preserve" Language="en-US">
    <DocumentSettings TopPage="0" DefaultTextStyle="0" DefaultLineStyle="0" DefaultFillStyle="0" DefaultGuideStyle="0"/>
    <Colors>
        <ColorEntry IX="0" RGB="#000000"/>
        <ColorEntry IX="1" RGB="#FFFFFF"/>
    </Colors>
    <FaceNames/>
    <StyleSheets/>
</VisioDocument>`;

export const DOCUMENT_RELS_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="${XML_NAMESPACES.RELATIONSHIPS}">
    <Relationship Id="rId1" Type="${RELATIONSHIP_TYPES.PAGES}" Target="pages/pages.xml"/>
    <Relationship Id="rId2" Type="${RELATIONSHIP_TYPES.WINDOWS}" Target="windows.xml"/>
</Relationships>`;

export const PAGES_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Pages xmlns="${XML_NAMESPACES.VISIO_MAIN}" xmlns:r="${XML_NAMESPACES.RELATIONSHIPS_OFFICE}" xml:space="preserve">
    <Page ID="1" Name="Page-1" NameU="Page-1" ViewScale="1" ViewCenterX="4.133858267716535" ViewCenterY="5.846456692913386"/>
</Pages>`;

export const PAGES_RELS_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="${XML_NAMESPACES.RELATIONSHIPS}">
    <Relationship Id="rId1" Type="${RELATIONSHIP_TYPES.PAGE}" Target="page1.xml"/>
</Relationships>`;

export const PAGE1_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<PageContents xmlns="${XML_NAMESPACES.VISIO_MAIN}" xmlns:r="${XML_NAMESPACES.RELATIONSHIPS_OFFICE}" xml:space="preserve">
    <Shapes/>
</PageContents>`;

export const WINDOWS_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Windows xmlns="${XML_NAMESPACES.VISIO_MAIN}" xmlns:r="${XML_NAMESPACES.RELATIONSHIPS_OFFICE}" xml:space="preserve">
    <Window ID="0" WindowType="Drawing" WindowState="1073741824" WindowLeft="0" WindowTop="0" WindowWidth="1024" WindowHeight="768" Container="0" ContainerType="Page" Sheet="0"/>
</Windows>`;

export const CORE_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="${XML_NAMESPACES.CORE_PROPERTIES}" xmlns:dc="${XML_NAMESPACES.DC_ELEMENTS}" xmlns:dcterms="${XML_NAMESPACES.DC_TERMS}" xmlns:dcmitype="${XML_NAMESPACES.DC_DCMITYPE}" xmlns:xsi="${XML_NAMESPACES.XSI}">
    <dc:title>Drawing1</dc:title>
    <dc:creator>ts-visio</dc:creator>
    <cp:lastModifiedBy>ts-visio</cp:lastModifiedBy>
    <cp:revision>1</cp:revision>
</cp:coreProperties>`;

export const APP_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="${XML_NAMESPACES.EXTENDED_PROPERTIES}" xmlns:vt="${XML_NAMESPACES.DOC_PROPS_VTYPES}">
    <Template>Basic</Template>
    <Application>ts-visio</Application>
</Properties>`;
