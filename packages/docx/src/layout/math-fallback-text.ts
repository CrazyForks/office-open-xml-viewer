import type { MathNode } from '@silurus/ooxml-core';
import type { DeepReadonly } from './types.js';

const SPACED_OPERATORS = new Set(['+', '-', '−', '=', '±', '×', '÷']);

function runText(text: string): string {
  return SPACED_OPERATORS.has(text) ? ` ${text} ` : text;
}

/** Parser-independent text fallback for one immutable OMML occurrence. */
export function mathFallbackText(nodes: readonly DeepReadonly<MathNode>[]): string {
  const renderNode = (node: DeepReadonly<MathNode>): string => {
    switch (node.kind) {
      case 'run': return runText(node.text);
      case 'fraction': return `${mathFallbackText(node.num)}/${mathFallbackText(node.den)}`;
      case 'sup': return `${mathFallbackText(node.base)}^${mathFallbackText(node.sup ?? [])}`;
      case 'sub': return `${mathFallbackText(node.base)}_${mathFallbackText(node.sub ?? [])}`;
      case 'subSup': return `${mathFallbackText(node.base)}_${mathFallbackText(node.sub ?? [])}^${mathFallbackText(node.sup ?? [])}`;
      case 'nary': return `${node.op}${mathFallbackText(node.sub ?? [])}${mathFallbackText(node.sup ?? [])}${mathFallbackText(node.body)}`;
      case 'delimiter': return `${node.begChar}${node.items.map(mathFallbackText).join(',')}${node.endChar}`;
      case 'radical': return `${node.index?.length ? mathFallbackText(node.index) : ''}√${mathFallbackText(node.radicand)}`;
      case 'limit': return `${mathFallbackText(node.base)}${mathFallbackText(node.lower ?? [])}${mathFallbackText(node.upper ?? [])}`;
      case 'array': return node.rows.map((row) => row.map(mathFallbackText).join(' ')).join(' ');
      case 'groupChr': return `${node.char}${mathFallbackText(node.base)}`;
      case 'bar':
      case 'box':
      case 'borderBox': return mathFallbackText(node.base);
      case 'accent': return `${node.char}${mathFallbackText(node.base)}`;
      case 'func': return `${mathFallbackText(node.name)}(${mathFallbackText(node.arg)})`;
      case 'group': return mathFallbackText(node.items);
      case 'phant': return node.show ? mathFallbackText(node.base) : '';
      case 'sPre': return `${mathFallbackText(node.sub)}${mathFallbackText(node.sup)}${mathFallbackText(node.base)}`;
    }
  };
  return nodes.map(renderNode).join('').replace(/[ \t]{2,}/g, ' ');
}
