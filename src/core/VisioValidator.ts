import { XMLParser } from 'fast-xml-parser';
import { VisioPackage } from '../VisioPackage';

export interface ValidationResult {
    valid: boolean;
    errors: string[];
    warnings: string[];
}

// Known Visio section names per Microsoft schema
const VALID_SECTION_NAMES = new Set([
    'Geometry', 'Character', 'Paragraph', 'Tabs', 'Scratch', 'Connection',
    'Field', 'Control', 'Action', 'Layer', 'Property', 'User', 'Hyperlink',
    'Reviewer', 'Annotation', 'ActionTag', 'Line', 'Fill', 'FillGradient',
    'LineGradient', 'TextXForm', 'RelQuadBezTo', 'RelCubBezTo', 'RelMoveTo',
    'RelLineTo', 'RelEllipticalArcTo', 'InfiniteLine', 'Ellipse', 'SplineStart',
    'SplineKnot', 'PolylineTo', 'NURBSTo'
]);

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

        // 5. Validate master references
        this.validateMasterReferences(pkg, errors, warnings);

        // 6. Validate relationship file integrity
        this.validateRelationshipIntegrity(pkg, errors, warnings);

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

            const overrides = parsed.Types.Override ?
                (Array.isArray(parsed.Types.Override) ? parsed.Types.Override : [parsed.Types.Override])
                : [];

            const partNames = overrides.map((o: any) => o['@_PartName']);

            if (!partNames.some((p: string) => p?.includes('document.xml'))) {
                warnings.push('[Content_Types].xml: Missing content type for document.xml');
            }

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

                const relId = pageNode.Rel?.['@_r:id'] || pageNode['@_r:id'];
                if (!relId) {
                    errors.push(`Page "${pageName}" (ID=${pageId}): Missing Rel element with r:id`);
                }

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

            this.validateShapeIds(parsed, pagePath, errors);
            this.validateConnects(parsed, pagePath, errors);
            this.validateSectionNames(parsed, pagePath, warnings);
            this.validateCellNames(parsed, pagePath, warnings);
            this.validateImageShapes(pkg, pageId, parsed, pagePath, errors, warnings);

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

    private validateSectionNames(parsed: any, pagePath: string, warnings: string[]): void {
        const shapes = this.getAllShapes(parsed);

        for (const shape of shapes) {
            if (!shape.Section) continue;

            const sections = Array.isArray(shape.Section) ? shape.Section : [shape.Section];
            for (const section of sections) {
                const name = section['@_N'];
                if (!name) {
                    warnings.push(`${pagePath}: Shape ${shape['@_ID']} has Section without N attribute`);
                }
            }
        }
    }

    private validateCellNames(parsed: any, pagePath: string, warnings: string[]): void {
        const shapes = this.getAllShapes(parsed);

        for (const shape of shapes) {
            // Check top-level cells
            if (shape.Cell) {
                const cells = Array.isArray(shape.Cell) ? shape.Cell : [shape.Cell];
                for (const cell of cells) {
                    if (!cell['@_N']) {
                        warnings.push(`${pagePath}: Shape ${shape['@_ID']} has Cell without N attribute`);
                    }
                }
            }

            // Check cells in sections
            if (shape.Section) {
                const sections = Array.isArray(shape.Section) ? shape.Section : [shape.Section];
                for (const section of sections) {
                    if (section.Cell) {
                        const cells = Array.isArray(section.Cell) ? section.Cell : [section.Cell];
                        for (const cell of cells) {
                            if (!cell['@_N']) {
                                warnings.push(`${pagePath}: Shape ${shape['@_ID']} Section ${section['@_N']} has Cell without N attribute`);
                            }
                        }
                    }
                    if (section.Row) {
                        const rows = Array.isArray(section.Row) ? section.Row : [section.Row];
                        for (const row of rows) {
                            if (row.Cell) {
                                const cells = Array.isArray(row.Cell) ? row.Cell : [row.Cell];
                                for (const cell of cells) {
                                    if (!cell['@_N']) {
                                        warnings.push(`${pagePath}: Shape ${shape['@_ID']} Row ${row['@_N'] || row['@_IX']} has Cell without N attribute`);
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }
    }

    private validateImageShapes(pkg: VisioPackage, pageId: string, parsed: any, pagePath: string, errors: string[], warnings: string[]): void {
        const shapes = this.getAllShapes(parsed);

        for (const shape of shapes) {
            if (shape['@_Type'] !== 'Foreign') continue;

            if (!shape.ForeignData) {
                warnings.push(`${pagePath}: Foreign shape ${shape['@_ID']} missing ForeignData`);
                continue;
            }

            const relId = shape.ForeignData.Rel?.['@_r:id'];
            if (relId) {
                try {
                    const relsPath = `visio/pages/_rels/page${pageId}.xml.rels`;
                    const relsContent = pkg.getFileText(relsPath);
                    const parsedRels = this.parser.parse(relsContent);

                    let rels = parsedRels.Relationships?.Relationship || [];
                    rels = Array.isArray(rels) ? rels : [rels];

                    const found = rels.find((r: any) => r['@_Id'] === relId);
                    if (!found) {
                        errors.push(`${pagePath}: Image shape ${shape['@_ID']} references non-existent relationship: ${relId}`);
                    }
                } catch {
                    // No rels file is okay if no relationships needed
                }
            }
        }
    }

    private validateMasterReferences(pkg: VisioPackage, errors: string[], warnings: string[]): void {
        const masterIds = new Set<string>();

        try {
            const mastersContent = pkg.getFileText('visio/masters/masters.xml');
            const parsedMasters = this.parser.parse(mastersContent);

            let masters = parsedMasters.Masters?.Master || [];
            masters = Array.isArray(masters) ? masters : [masters];

            for (const master of masters) {
                if (master['@_ID']) {
                    masterIds.add(master['@_ID']);
                }
            }
        } catch {
            return; // No masters.xml is okay
        }

        try {
            const pagesContent = pkg.getFileText('visio/pages/pages.xml');
            const parsedPages = this.parser.parse(pagesContent);

            let pageNodes = parsedPages.Pages?.Page || [];
            pageNodes = Array.isArray(pageNodes) ? pageNodes : [pageNodes];

            for (const pageNode of pageNodes) {
                const pageId = pageNode['@_ID'];
                const pagePath = `visio/pages/page${pageId}.xml`;

                try {
                    const pageContent = pkg.getFileText(pagePath);
                    const parsed = this.parser.parse(pageContent);
                    const shapes = this.getAllShapes(parsed);

                    for (const shape of shapes) {
                        const masterId = shape['@_Master'];
                        if (masterId && !masterIds.has(masterId)) {
                            errors.push(`${pagePath}: Shape ${shape['@_ID']} references non-existent Master: ${masterId}`);
                        }
                    }
                } catch {
                    // Page file issues handled elsewhere
                }
            }
        } catch {
            // Pages issues handled elsewhere
        }
    }

    private validateRelationshipIntegrity(pkg: VisioPackage, errors: string[], warnings: string[]): void {
        try {
            const relsContent = pkg.getFileText('visio/pages/_rels/pages.xml.rels');
            const parsedRels = this.parser.parse(relsContent);

            let rels = parsedRels.Relationships?.Relationship || [];
            rels = Array.isArray(rels) ? rels : [rels];

            for (const rel of rels) {
                const target = rel['@_Target'];
                const id = rel['@_Id'];

                if (!id) {
                    errors.push('pages.xml.rels: Relationship missing Id attribute');
                    continue;
                }

                if (!target) {
                    errors.push(`pages.xml.rels: Relationship ${id} missing Target attribute`);
                    continue;
                }

                const fullPath = `visio/pages/${target}`;
                try {
                    pkg.getFileText(fullPath);
                } catch {
                    errors.push(`pages.xml.rels: Relationship ${id} targets non-existent file: ${target}`);
                }
            }
        } catch {
            // Rels file issues might be handled elsewhere
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
