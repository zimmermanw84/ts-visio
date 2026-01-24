import { XMLParser } from 'fast-xml-parser';
import { VisioPackage } from '../VisioPackage';

export interface ValidationResult {
    valid: boolean;
    errors: string[];
    warnings: string[];
}

export class VisioValidator {
    private parser: XMLParser;

    constructor() {
        this.parser = new XMLParser({
            ignoreAttributes: false,
            attributeNamePrefix: '@_'
        });
    }

    /**
     * Validate a VisioPackage for structural correctness
     */
    async validate(pkg: VisioPackage): Promise<ValidationResult> {
        const errors: string[] = [];
        const warnings: string[] = [];

        // 1. Check required files exist
        this.checkRequiredFiles(pkg, errors);

        // 2. Validate Content Types
        this.validateContentTypes(pkg, errors, warnings);

        // 3. Validate document relationships
        this.validateDocumentRels(pkg, errors);

        // 4. Validate pages
        this.validatePages(pkg, errors, warnings);

        return {
            valid: errors.length === 0,
            errors,
            warnings
        };
    }

    private checkRequiredFiles(pkg: VisioPackage, errors: string[]): void {
        const requiredFiles = [
            '[Content_Types].xml',
            '_rels/.rels',
            'visio/document.xml',
            'visio/_rels/document.xml.rels',
            'visio/pages/pages.xml',
            'visio/pages/_rels/pages.xml.rels'
        ];

        for (const file of requiredFiles) {
            try {
                pkg.getFileText(file);
            } catch {
                errors.push(`Missing required file: ${file}`);
            }
        }
    }

    private validateContentTypes(pkg: VisioPackage, errors: string[], warnings: string[]): void {
        try {
            const content = pkg.getFileText('[Content_Types].xml');
            const parsed = this.parser.parse(content);

            if (!parsed.Types) {
                errors.push('[Content_Types].xml: Missing root <Types> element');
                return;
            }

            // Check for required content types
            const overrides = parsed.Types.Override ?
                (Array.isArray(parsed.Types.Override) ? parsed.Types.Override : [parsed.Types.Override])
                : [];

            const partNames = overrides.map((o: any) => o['@_PartName']);

            // Check document.xml is registered
            if (!partNames.some((p: string) => p?.includes('document.xml'))) {
                warnings.push('[Content_Types].xml: Missing content type for document.xml');
            }

            // Check pages.xml is registered
            if (!partNames.some((p: string) => p?.includes('pages.xml'))) {
                warnings.push('[Content_Types].xml: Missing content type for pages.xml');
            }
        } catch (e: any) {
            errors.push(`[Content_Types].xml: ${e.message}`);
        }
    }

    private validateDocumentRels(pkg: VisioPackage, errors: string[]): void {
        try {
            const content = pkg.getFileText('visio/_rels/document.xml.rels');
            const parsed = this.parser.parse(content);

            if (!parsed.Relationships) {
                errors.push('document.xml.rels: Missing <Relationships> element');
                return;
            }

            const rels = parsed.Relationships.Relationship ?
                (Array.isArray(parsed.Relationships.Relationship) ? parsed.Relationships.Relationship : [parsed.Relationships.Relationship])
                : [];

            // Check pages relationship exists
            const pagesRel = rels.find((r: any) => r['@_Target']?.includes('pages'));
            if (!pagesRel) {
                errors.push('document.xml.rels: Missing relationship to pages.xml');
            }
        } catch (e: any) {
            errors.push(`document.xml.rels: ${e.message}`);
        }
    }

    private validatePages(pkg: VisioPackage, errors: string[], warnings: string[]): void {
        try {
            const pagesContent = pkg.getFileText('visio/pages/pages.xml');
            const parsedPages = this.parser.parse(pagesContent);

            let pageNodes = parsedPages.Pages?.Page;
            if (!pageNodes) {
                warnings.push('pages.xml: No pages defined');
                return;
            }

            pageNodes = Array.isArray(pageNodes) ? pageNodes : [pageNodes];

            for (const pageNode of pageNodes) {
                const pageId = pageNode['@_ID'];
                const pageName = pageNode['@_Name'] || `Page-${pageId}`;

                if (!pageId) {
                    errors.push(`Page "${pageName}": Missing ID attribute`);
                    continue;
                }

                // Check for Rel child element with r:id (required per MS schema)
                const relId = pageNode.Rel?.['@_r:id'] || pageNode['@_r:id'];
                if (!relId) {
                    errors.push(`Page "${pageName}" (ID=${pageId}): Missing Rel element with r:id`);
                }

                // Validate page content file exists (page{ID}.xml)
                this.validatePageContent(pkg, pageId, pageName, errors, warnings);
            }
        } catch (e: any) {
            errors.push(`pages.xml: ${e.message}`);
        }
    }

    private validatePageContent(pkg: VisioPackage, pageId: string, pageName: string, errors: string[], warnings: string[]): void {
        const pagePath = `visio/pages/page${pageId}.xml`;

        try {
            const content = pkg.getFileText(pagePath);
            const parsed = this.parser.parse(content);

            if (!parsed.PageContents) {
                errors.push(`${pagePath}: Missing <PageContents> element`);
                return;
            }

            // Validate shapes have unique IDs
            this.validateShapeIds(parsed, pagePath, errors);

            // Validate connectors reference existing shapes
            this.validateConnects(parsed, pagePath, errors);

        } catch (e: any) {
            errors.push(`${pagePath}: ${e.message}`);
        }
    }

    private validateShapeIds(parsed: any, pagePath: string, errors: string[]): void {
        const shapes = this.getAllShapes(parsed);
        const ids = new Set<string>();

        for (const shape of shapes) {
            const id = shape['@_ID'];
            if (!id) {
                errors.push(`${pagePath}: Shape missing ID attribute`);
                continue;
            }
            if (ids.has(id)) {
                errors.push(`${pagePath}: Duplicate shape ID: ${id}`);
            }
            ids.add(id);
        }
    }

    private validateConnects(parsed: any, pagePath: string, errors: string[]): void {
        if (!parsed.PageContents.Connects?.Connect) return;

        const shapes = this.getAllShapes(parsed);
        const shapeIds = new Set(shapes.map((s: any) => s['@_ID']));

        let connects = parsed.PageContents.Connects.Connect;
        connects = Array.isArray(connects) ? connects : [connects];

        for (const connect of connects) {
            const fromSheet = connect['@_FromSheet'];
            const toSheet = connect['@_ToSheet'];

            if (fromSheet && !shapeIds.has(fromSheet)) {
                errors.push(`${pagePath}: Connect references non-existent FromSheet: ${fromSheet}`);
            }
            if (toSheet && !shapeIds.has(toSheet)) {
                errors.push(`${pagePath}: Connect references non-existent ToSheet: ${toSheet}`);
            }
        }
    }

    private getAllShapes(parsed: any): any[] {
        let topLevelShapes = parsed.PageContents.Shapes?.Shape || [];
        if (!Array.isArray(topLevelShapes)) {
            topLevelShapes = topLevelShapes ? [topLevelShapes] : [];
        }

        const all: any[] = [];
        const gather = (shapeList: any[]): void => {
            for (const s of shapeList) {
                all.push(s);
                if (s.Shapes?.Shape) {
                    const children = Array.isArray(s.Shapes.Shape) ? s.Shapes.Shape : [s.Shapes.Shape];
                    gather(children);
                }
            }
        };

        gather(topLevelShapes);
        return all;
    }
}
