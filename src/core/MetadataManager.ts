import { VisioPackage } from '../VisioPackage';
import { DocumentMetadata } from '../types/VisioTypes';
import { XML_NAMESPACES } from './VisioConstants';
import { createXmlParser, createXmlBuilder, buildXml } from '../utils/XmlHelper';

export class MetadataManager {
    private parser = createXmlParser();
    private builder = createXmlBuilder();

    constructor(private pkg: VisioPackage) {}

    // ---- public API --------------------------------------------------------

    /** Read document metadata from `docProps/core.xml` and `docProps/app.xml`. */
    getMetadata(): DocumentMetadata {
        const core = this.parsedCore();
        const app  = this.parsedApp();

        return {
            title:          this.str(core['dc:title']),
            author:         this.str(core['dc:creator']),
            description:    this.str(core['dc:description']),
            keywords:       this.str(core['cp:keywords']),
            lastModifiedBy: this.str(core['cp:lastModifiedBy']),
            created:        this.date(core['dcterms:created']),
            modified:       this.date(core['dcterms:modified']),
            company:        this.str(app['Company']),
            manager:        this.str(app['Manager']),
        };
    }

    /** Merge the supplied fields into the existing metadata and persist to the package. */
    setMetadata(props: Partial<DocumentMetadata>): void {
        const coreKeys: (keyof DocumentMetadata)[] = [
            'title', 'author', 'description', 'keywords', 'lastModifiedBy', 'created', 'modified',
        ];
        const appKeys: (keyof DocumentMetadata)[] = ['company', 'manager'];

        if (coreKeys.some(k => k in props)) this.writeCore(props);
        if (appKeys.some(k => k in props))  this.writeApp(props);
    }

    // ---- private helpers ---------------------------------------------------

    /** Extract a string from a parsed XML node (handles plain strings and #text objects). */
    private str(val: unknown): string | undefined {
        if (val === undefined || val === null) return undefined;
        if (typeof val === 'string') return val || undefined;
        if (typeof val === 'object' && '#text' in (val as Record<string, unknown>)) {
            const t = (val as Record<string, unknown>)['#text'];
            const s = typeof t === 'string' ? t : String(t);
            return s || undefined;
        }
        return undefined;
    }

    /** Parse an ISO datetime string into a Date (returns undefined on failure). */
    private date(val: unknown): Date | undefined {
        const text = this.str(val);
        if (!text) return undefined;
        const d = new Date(text);
        return isNaN(d.getTime()) ? undefined : d;
    }

    private parsedCore(): Record<string, unknown> {
        try {
            const xml = this.pkg.getFileText('docProps/core.xml');
            const parsed = this.parser.parse(xml) as Record<string, unknown>;
            return (parsed['cp:coreProperties'] as Record<string, unknown>) ?? {};
        } catch {
            return {};
        }
    }

    private parsedApp(): Record<string, unknown> {
        try {
            const xml = this.pkg.getFileText('docProps/app.xml');
            const parsed = this.parser.parse(xml) as Record<string, unknown>;
            return (parsed['Properties'] as Record<string, unknown>) ?? {};
        } catch {
            return {};
        }
    }

    private writeCore(props: Partial<DocumentMetadata>): void {
        let parsed: Record<string, unknown>;
        try {
            const xml = this.pkg.getFileText('docProps/core.xml');
            parsed = this.parser.parse(xml) as Record<string, unknown>;
        } catch {
            parsed = this.blankCoreParsed();
        }

        const root = parsed['cp:coreProperties'] as Record<string, unknown>;

        if ('title'          in props) root['dc:title']          = props.title ?? '';
        if ('author'         in props) root['dc:creator']        = props.author ?? '';
        if ('description'    in props) root['dc:description']    = props.description ?? '';
        if ('keywords'       in props) root['cp:keywords']       = props.keywords ?? '';
        if ('lastModifiedBy' in props) root['cp:lastModifiedBy'] = props.lastModifiedBy ?? '';

        if ('created' in props && props.created !== undefined) {
            root['dcterms:created'] = {
                '@_xsi:type': 'dcterms:W3CDTF',
                '#text': props.created.toISOString(),
            };
        }
        if ('modified' in props && props.modified !== undefined) {
            root['dcterms:modified'] = {
                '@_xsi:type': 'dcterms:W3CDTF',
                '#text': props.modified.toISOString(),
            };
        }

        this.pkg.updateFile('docProps/core.xml', buildXml(this.builder, parsed));
    }

    private writeApp(props: Partial<DocumentMetadata>): void {
        let parsed: Record<string, unknown>;
        try {
            const xml = this.pkg.getFileText('docProps/app.xml');
            parsed = this.parser.parse(xml) as Record<string, unknown>;
        } catch {
            parsed = this.blankAppParsed();
        }

        const root = parsed['Properties'] as Record<string, unknown>;
        if ('company' in props) root['Company'] = props.company ?? '';
        if ('manager' in props) root['Manager'] = props.manager ?? '';

        this.pkg.updateFile('docProps/app.xml', buildXml(this.builder, parsed));
    }

    private blankCoreParsed(): Record<string, unknown> {
        return {
            'cp:coreProperties': {
                '@_xmlns:cp':      XML_NAMESPACES.CORE_PROPERTIES,
                '@_xmlns:dc':      XML_NAMESPACES.DC_ELEMENTS,
                '@_xmlns:dcterms': XML_NAMESPACES.DC_TERMS,
                '@_xmlns:dcmitype': XML_NAMESPACES.DC_DCMITYPE,
                '@_xmlns:xsi':     XML_NAMESPACES.XSI,
            },
        };
    }

    private blankAppParsed(): Record<string, unknown> {
        return {
            'Properties': {
                '@_xmlns':    XML_NAMESPACES.EXTENDED_PROPERTIES,
                '@_xmlns:vt': XML_NAMESPACES.DOC_PROPS_VTYPES,
                'Application': 'ts-visio',
                'Template': 'Basic',
            },
        };
    }
}
