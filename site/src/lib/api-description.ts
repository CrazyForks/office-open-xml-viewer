export interface DescribedItem {
  desc: string;
  emphasis?: string;
}

export interface DescriptionToken {
  text: string;
  code: boolean;
  emphasized: boolean;
}

function inlineCodeTokens(text: string, emphasized: boolean): DescriptionToken[] {
  const parts = text.split('`');
  if (parts.length % 2 === 0) {
    throw new Error(`Unmatched backtick in API description: ${text}`);
  }
  return parts
    .map((part, index) => ({ text: part, code: index % 2 === 1, emphasized }))
    .filter(({ text: part }) => part.length > 0);
}

export function descriptionTokens(item: DescribedItem): DescriptionToken[] {
  if (!item.emphasis) return inlineCodeTokens(item.desc, false);

  const start = item.desc.indexOf(item.emphasis);
  if (start < 0) {
    throw new Error(`API emphasis must be an exact description substring: ${item.emphasis}`);
  }

  return [
    ...inlineCodeTokens(item.desc.slice(0, start), false),
    ...inlineCodeTokens(item.emphasis, true),
    ...inlineCodeTokens(item.desc.slice(start + item.emphasis.length), false),
  ];
}
