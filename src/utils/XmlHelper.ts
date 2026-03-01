import { XMLParser, XMLBuilder } from 'fast-xml-parser';

const XML_DECLARATION = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';

/**
 * Standard XMLParser options used across all Visio XML parts.
 *
 * - ignoreDeclaration: false  → preserves <?xml?> in the parsed object
 * - ignoreAttributes: false   → preserves xmlns, xmlns:r, xml:space, etc.
 * - parseAttributeValue: false → keeps attribute values as strings (avoids
 *   numeric coercion on IDs, versions, etc.)
 */
export const PARSER_OPTIONS = {
    ignoreAttributes: false,
    attributeNamePrefix: '@_',
    ignoreDeclaration: false,
    parseAttributeValue: false,
} as const;

/**
 * Standard XMLBuilder options used across all Visio XML parts.
 *
 * - suppressBooleanAttributes: false → always emit attribute="value", never
 *   bare attribute (important for Visio attributes like xml:space="preserve")
 */
export const BUILDER_OPTIONS = {
    ignoreAttributes: false,
    attributeNamePrefix: '@_',
    format: true,
    suppressBooleanAttributes: false,
} as const;

export function createXmlParser(): XMLParser {
    return new XMLParser(PARSER_OPTIONS);
}

export function createXmlBuilder(): XMLBuilder {
    return new XMLBuilder(BUILDER_OPTIONS);
}

/**
 * Serialize a parsed XML object to a string, guaranteeing the XML declaration
 * is always present at the top of the output.
 *
 * fast-xml-parser's XMLBuilder re-emits the `?xml` key when it is present in
 * the parsed object (requires ignoreDeclaration: false on the parser side).
 * This function acts as a safety net for cases where the declaration was
 * absent in the source XML or was stripped for any reason.
 */
export function buildXml(builder: XMLBuilder, parsed: unknown): string {
    const xml = builder.build(parsed);
    if (xml.trimStart().startsWith('<?xml')) {
        return xml;
    }
    return XML_DECLARATION + '\n' + xml;
}
