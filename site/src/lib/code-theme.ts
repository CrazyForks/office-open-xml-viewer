/**
 * A restrained Vitesse-inspired pair built around the site's fluorescent lime.
 * The palette deliberately stays inside a botanical range, with clay reserved
 * for literals/errors and blue-green reserved for types.
 */
const scopes = {
  comments: ['comment', 'punctuation.definition.comment', 'string.comment'],
  punctuation: [
    'delimiter',
    'delimiter.bracket',
    'meta.brace',
    'meta.type.annotation',
    'punctuation',
    'storage.type.function.arrow',
  ],
  keywords: [
    'keyword',
    'storage.modifier',
    'storage.type.class.jsdoc',
    'punctuation.definition.template-expression',
  ],
  storage: [
    'storage.type',
    'support.type.builtin',
    'constant.language.undefined',
    'constant.language.null',
  ],
  functions: ['entity.name.function', 'entity.name.method', 'support.function'],
  strings: ['string', 'attribute.value'],
  properties: [
    'property',
    'meta.property-name',
    'meta.object-literal.key',
    'attribute.name',
    'support.type.property-name',
  ],
  variables: ['variable', 'identifier', 'entity.other.attribute-name'],
  types: ['support.type.primitive', 'entity.name.type', 'type.identifier', 'support.class'],
  numbers: ['constant.numeric', 'number', 'support.constant'],
  constants: ['constant.language', 'entity.name.constant', 'variable.language'],
  tags: ['entity.name.tag', 'tag.html'],
};

const limeLight = {
  name: 'ooxml-lime-light',
  displayName: 'OOXML Lime Light',
  type: 'light' as const,
  colors: {
    'editor.background': '#f9f8f4',
    'editor.foreground': '#2c2d2a',
    'editor.selectionBackground': '#c9ff4340',
    'editor.lineHighlightBackground': '#ecebe6',
  },
  semanticHighlighting: true,
  semanticTokenColors: {
    class: '#356c60',
    interface: '#356c60',
    type: '#356c60',
    property: '#58611e',
    function: '#52751b',
  },
  tokenColors: [
    { settings: { foreground: '#2c2d2a', background: '#f9f8f4' } },
    { scope: scopes.comments, settings: { foreground: '#70726c', fontStyle: 'italic' } },
    { scope: scopes.punctuation, settings: { foreground: '#8b8c85' } },
    { scope: scopes.keywords, settings: { foreground: '#376241' } },
    { scope: scopes.storage, settings: { foreground: '#8a5738' } },
    { scope: scopes.functions, settings: { foreground: '#52751b' } },
    { scope: scopes.strings, settings: { foreground: '#5f7428' } },
    { scope: scopes.properties, settings: { foreground: '#58611e' } },
    { scope: scopes.variables, settings: { foreground: '#71592f' } },
    { scope: scopes.types, settings: { foreground: '#356c60' } },
    { scope: scopes.numbers, settings: { foreground: '#75641d' } },
    { scope: scopes.constants, settings: { foreground: '#4f761f' } },
    { scope: scopes.tags, settings: { foreground: '#376241' } },
    { scope: ['keyword.operator', 'meta.var.expr.ts'], settings: { foreground: '#487137' } },
    { scope: ['invalid', 'message.error'], settings: { foreground: '#b9563e' } },
    { scope: ['markup.heading', 'markup.raw'], settings: { foreground: '#416300', fontStyle: 'bold' } },
  ],
};

const limeDark = {
  name: 'ooxml-lime-dark',
  displayName: 'OOXML Lime Dark',
  type: 'dark' as const,
  colors: {
    'editor.background': '#191a17',
    'editor.foreground': '#e1e4dc',
    'editor.selectionBackground': '#c9ff4330',
    'editor.lineHighlightBackground': '#23251f',
  },
  semanticHighlighting: true,
  semanticTokenColors: {
    class: '#79b9a5',
    interface: '#79b9a5',
    type: '#79b9a5',
    property: '#d0d76d',
    function: '#a8ce60',
  },
  tokenColors: [
    { settings: { foreground: '#e1e4dc', background: '#191a17' } },
    { scope: scopes.comments, settings: { foreground: '#93988e', fontStyle: 'italic' } },
    { scope: scopes.punctuation, settings: { foreground: '#858980' } },
    { scope: scopes.keywords, settings: { foreground: '#c9ff43' } },
    { scope: scopes.storage, settings: { foreground: '#e29a72' } },
    { scope: scopes.functions, settings: { foreground: '#a8ce60' } },
    { scope: scopes.strings, settings: { foreground: '#9fc57a' } },
    { scope: scopes.properties, settings: { foreground: '#d0d76d' } },
    { scope: scopes.variables, settings: { foreground: '#d1b57a' } },
    { scope: scopes.types, settings: { foreground: '#79b9a5' } },
    { scope: scopes.numbers, settings: { foreground: '#d5c26d' } },
    { scope: scopes.constants, settings: { foreground: '#b3d45a' } },
    { scope: scopes.tags, settings: { foreground: '#c9ff43' } },
    { scope: ['keyword.operator', 'meta.var.expr.ts'], settings: { foreground: '#b6dc6e' } },
    { scope: ['invalid', 'message.error'], settings: { foreground: '#ec9077' } },
    { scope: ['markup.heading', 'markup.raw'], settings: { foreground: '#c9ff43', fontStyle: 'bold' } },
  ],
};

/** Shared Shiki pair so every public code sample follows the site theme. */
export const codeThemes = {
  light: limeLight,
  dark: limeDark,
};
