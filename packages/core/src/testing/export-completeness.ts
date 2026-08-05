/**
 * Export-completeness checker (v1.0 API freeze tooling).
 *
 * Uses the TypeScript Compiler API to find public types that are *reachable*
 * from a package's `index.ts` barrel but are NOT themselves re-exported. The
 * classic failure this guards against: a union like
 *
 *   export type SlideElement = ShapeElement | TableElement | ChartElement;
 *
 * is exported, but `TableElement` / `ChartElement` are not — so a consumer can
 * receive a `SlideElement`, narrow on `el.type === 'table'`, and have no name
 * for the resulting object.
 *
 * The traversal is driven by the *type checker* (not the raw AST), so it
 * follows only the public type surface:
 *   - union / intersection constituents,
 *   - type arguments (e.g. `Array<T>`, `Map<K, V>`, `T | null`),
 *   - the **public** properties of object types (private / protected class
 *     members and function-body-local type aliases are invisible to the type
 *     API and are therefore correctly skipped),
 *   - call / construct signature parameter and return types.
 *
 * It is intentionally conservative: only types *declared inside the same
 * package* (under the package's `src/` dir) are required to be exported. Types
 * from `@silurus/ooxml-core` or `lib.dom` (`HTMLCanvasElement`, `Error`, …) are
 * out of scope.
 *
 * This is test-only tooling and is never re-exported from the package barrels,
 * so it does not enter the published bundle.
 */
import {
  API,
  SignatureKind,
  SymbolFlags,
  type Symbol as TypeScriptSymbol,
  type Type,
} from 'typescript/unstable/sync';
import {
  ModifierFlags,
  type Declaration,
  type Node,
} from 'typescript/unstable/ast';
import {
  isClassDeclaration,
  isEnumDeclaration,
  isInterfaceDeclaration,
  isPrivateIdentifier,
  isTypeAliasDeclaration,
} from 'typescript/unstable/ast/is';
import * as path from 'node:path';
import { fileURLToPath } from 'node:url';

export interface MissingExport {
  /** The unexported type's name. */
  name: string;
  /** Absolute path of the file that declares it. */
  declaredIn: string;
  /** The exported barrel symbol through which it is reachable (diagnostics). */
  reachableFrom: string[];
}

export interface CheckOptions {
  /** Absolute path to the package's `src/index.ts` barrel. */
  indexPath: string;
  /**
   * Absolute path to the package `src/` directory. Only types declared under
   * this directory are considered "same-package" and therefore required to be
   * exported. Defaults to `dirname(indexPath)`.
   */
  srcDir?: string;
  /**
   * Names to ignore even if reachable and unexported (escape hatch for
   * deliberately-internal helper types). Usually empty.
   */
  allowlist?: readonly string[];
}

/** Normalise a fs path for cross-platform prefix comparison. */
function norm(p: string): string {
  return path.normalize(p).split(path.sep).join('/');
}

/** True when `file` lives under `dir` (both already normalised). */
function isUnder(file: string, dir: string): boolean {
  const f = norm(file);
  const d = norm(dir).replace(/\/$/, '');
  return f === d || f.startsWith(d + '/');
}

/**
 * Find types reachable from the barrel's exports that are declared in the same
 * package's `src/` tree but are not themselves exported from the barrel.
 */
export function findMissingExports(opts: CheckOptions): MissingExport[] {
  const indexPath = path.resolve(opts.indexPath);
  const srcDir = norm(opts.srcDir ?? path.dirname(opts.indexPath));
  const allow = new Set(opts.allowlist ?? []);
  const api = new API({ cwd: path.dirname(indexPath) });
  const snapshot = api.updateSnapshot({ openFiles: [indexPath] });

  try {
    const project = snapshot.getDefaultProjectForFile(indexPath);
    if (!project) throw new Error(`Could not resolve TypeScript project for: ${indexPath}`);
    const program = project.program;
    const checker = project.checker;

    const indexSource = program.getSourceFile(indexPath);
    if (!indexSource) throw new Error(`Could not load index source file: ${indexPath}`);
    const moduleSymbol = checker.getSymbolAtLocation(indexSource);
    if (!moduleSymbol) throw new Error(`Could not resolve module symbol for: ${indexPath}`);

    const exports = checker.getExportsOfModule(moduleSymbol);
    const exportedNames = new Set<string>(exports.map((symbol) => symbol.name));

    const missing = new Map<string, MissingExport>();
    // TypeScript 7 exposes stable type IDs through its native sync API.
    const visitedTypeIds = new Set<number>();
    const MAX_DEPTH = 64;

    const resolveDeclaration = (
      handle: TypeScriptSymbol['valueDeclaration'],
    ): Declaration | undefined => handle?.resolve(project) as Declaration | undefined;

    const declarationsOf = (symbol: TypeScriptSymbol): Declaration[] =>
      symbol.declarations
        .map((handle) => handle.resolve(project) as Declaration | undefined)
        .filter((declaration): declaration is Declaration => declaration !== undefined);

    function resolveAlias(symbol: TypeScriptSymbol): TypeScriptSymbol {
      return symbol.flags & SymbolFlags.Alias ? checker.getAliasedSymbol(symbol) : symbol;
    }

    /** Where is a type's naming symbol declared? */
    type Origin =
      | { kind: 'in-package'; name: string; file: string }
      | { kind: 'external' } // named, but declared outside the package src (core, lib.dom, …)
      | { kind: 'anonymous' }; // inline object literal / union / primitive — no naming symbol

    /**
     * True for class members that are not part of the public surface: explicit
     * `private`/`protected` modifiers, ECMAScript `#private` names, or
     * `@internal`-style underscore-prefixed members are NOT treated as private
     * here (only real access modifiers / `#` names are). The Compiler API leaks
     * private members through `getProperties()`, so we must filter them.
     */
    function isNonPublicMember(symbol: TypeScriptSymbol): boolean {
      for (const declaration of declarationsOf(symbol)) {
        const modifierFlags = (
          declaration as Declaration & { readonly modifierFlags?: ModifierFlags }
        ).modifierFlags ?? ModifierFlags.None;
        if (
          modifierFlags &
          (ModifierFlags.Private | ModifierFlags.Protected)
        ) {
          return true;
        }
        // `#private` fields/methods.
        const name = (declaration as Declaration & { readonly name?: Node }).name;
        if (name && isPrivateIdentifier(name)) return true;
      }
      return false;
    }

    /** Only these declaration kinds introduce a *named type* in the API surface. */
    function isTypeDeclaration(declaration: Declaration): boolean {
      return (
        isInterfaceDeclaration(declaration) ||
        isTypeAliasDeclaration(declaration) ||
        isClassDeclaration(declaration) ||
        isEnumDeclaration(declaration)
      );
    }

    function originOf(type: Type): Origin {
      const symbol = type.getAliasSymbol() ?? type.getSymbol();
      if (!symbol || symbol.flags & SymbolFlags.TypeParameter) return { kind: 'anonymous' };
      const name = symbol.name;
      if (!name || name === '__type' || name === '__object') return { kind: 'anonymous' };
      const declarations = declarationsOf(symbol);
      if (declarations.length === 0) return { kind: 'external' };
      // A symbol whose declarations are functions/methods/variables (not a type
      // declaration) does not name a *type* — it is an anonymous structural type
      // for our purposes (we still descend into its signature via the caller).
      const typeDeclarations = declarations.filter(isTypeDeclaration);
      if (typeDeclarations.length === 0) return { kind: 'anonymous' };
      for (const declaration of typeDeclarations) {
        const file = declaration.getSourceFile().fileName;
        if (isUnder(file, srcDir)) return { kind: 'in-package', name, file: norm(file) };
      }
      return { kind: 'external' };
    }

    function record(name: string, file: string, root: string): void {
      const existing = missing.get(name);
      if (existing) {
        if (!existing.reachableFrom.includes(root)) existing.reachableFrom.push(root);
      } else {
        missing.set(name, { name, declaredIn: file, reachableFrom: [root] });
      }
    }

    function typeArgumentsOf(type: Type): readonly Type[] {
      const typeArguments = [...type.getAliasTypeArguments()];
      if (type.isTypeReference()) typeArguments.push(...checker.getTypeArguments(type));
      return Array.from(new Map(typeArguments.map((argument) => [argument.id, argument])).values());
    }

    /** Recurse through a type's public surface, recording in-package types. */
    function walkType(type: Type, root: string, depth: number): void {
      if (depth > MAX_DEPTH || visitedTypeIds.has(type.id)) return;
      visitedTypeIds.add(type.id);

      const origin = originOf(type);

      // External named type (core, lib.dom, Node, …): record nothing and DO NOT
      // descend into its members. This is the critical pruning step — without it
      // the walk would explore the entire DOM/lib type graph and OOM.
      if (origin.kind === 'external') {
        // Type arguments still matter: an external container like `Array<Foo>` or
        // `Promise<Foo>` may carry an in-package `Foo`. Descend only into those.
        for (const argument of typeArgumentsOf(type)) walkType(argument, root, depth + 1);
        return;
      }

      // In-package named type: check + record if unexported, then keep walking
      // its structure (it may reach further in-package types).
      if (origin.kind === 'in-package') {
        if (!exportedNames.has(origin.name) && !allow.has(origin.name)) {
          record(origin.name, origin.file, root);
        }
      }

      // Union / intersection constituents.
      if (type.isUnionType() || type.isIntersectionType()) {
        for (const member of type.getTypes()) walkType(member, root, depth + 1);
      }

      // Type arguments (Array<T>, Map<K,V>, generics on in-package aliases …).
      for (const argument of typeArgumentsOf(type)) walkType(argument, root, depth + 1);

      // Own properties — only descend for in-package and anonymous (inline
      // object-literal) types. NB: the Compiler API's `getProperties()` DOES
      // include `private`/`protected` class members (privacy is enforced at
      // check time, not stripped from the symbol table), so they must be
      // filtered out explicitly — otherwise the walk leaks through a viewer's
      // private worker bridge into the internal message protocol types.
      for (const property of checker.getPropertiesOfType(type)) {
        if (isNonPublicMember(property)) continue;
        const declaration =
          resolveDeclaration(property.valueDeclaration) ?? declarationsOf(property)[0];
        if (!declaration) continue;
        walkType(checker.getTypeOfSymbolAtLocation(property, declaration), root, depth + 1);
      }

      // Call / construct signatures: parameter and return types.
      const signatures = [
        ...checker.getSignaturesOfType(type, SignatureKind.Call),
        ...checker.getSignaturesOfType(type, SignatureKind.Construct),
      ];
      for (const signature of signatures) {
        for (const parameter of signature.getParameters()) {
          const declaration =
            resolveDeclaration(parameter.valueDeclaration) ?? declarationsOf(parameter)[0];
          if (!declaration) continue;
          walkType(checker.getTypeOfSymbolAtLocation(parameter, declaration), root, depth + 1);
        }
        const returnType = checker.getReturnTypeOfSignature(signature);
        if (returnType) walkType(returnType, root, depth + 1);
      }
    }

    for (const exported of exports) {
      const symbol = resolveAlias(exported);
      const declaration = declarationsOf(symbol)[0];
      if (!declaration) continue;
      // Use the declared type at its declaration site so type aliases resolve to
      // their target (unions, object literals, …) and classes/interfaces to their
      // instance type.
      const type = checker.getTypeOfSymbolAtLocation(symbol, declaration);
      walkType(type, exported.name, 0);
      // For interfaces / type aliases the symbol type above may be the *type* of
      // a value; also walk the declared type to be safe.
      walkType(checker.getDeclaredTypeOfSymbol(symbol), exported.name, 0);
    }

    return Array.from(missing.values()).sort((a, b) => a.name.localeCompare(b.name));
  } finally {
    snapshot.dispose();
    api.close();
  }
}

/**
 * Convenience wrapper for package tests: resolve `index.ts` relative to the
 * test module's `import.meta.url` so callers don't need `@silurus/ooxml-core`'s
 * `@types/node` to call `fileURLToPath` themselves. The node-API dependency
 * lives here in core, which already depends on `@types/node`.
 *
 * @param metaUrl     the test file's `import.meta.url`.
 * @param relIndexPath path to the barrel relative to the test file (default
 *                     `./index.ts`).
 */
export function findMissingExportsFromUrl(
  metaUrl: string,
  relIndexPath = './index.ts',
  extra?: Omit<CheckOptions, 'indexPath'>,
): MissingExport[] {
  const indexPath = fileURLToPath(new URL(relIndexPath, metaUrl));
  return findMissingExports({ indexPath, ...extra });
}

/**
 * Convenience: format a readable failure message listing the reachable-but-
 * unexported types, or '' on success.
 */
export function formatMissing(missing: MissingExport[]): string {
  if (missing.length === 0) return '';
  const lines = missing.map(
    (m) =>
      `  - ${m.name} (declared in ${path.basename(m.declaredIn)}; reachable from ${m.reachableFrom.join(', ')})`,
  );
  return `Reachable-but-unexported public types:\n${lines.join('\n')}`;
}
