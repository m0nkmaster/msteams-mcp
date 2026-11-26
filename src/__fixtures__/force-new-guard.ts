import ts from 'typescript';

/**
 * Test helper: true when a piece of advice recommends forceNew without
 * requiring user approval.
 *
 * Mentions are forceNew/force_new/force-new or the CLI --force flag. A
 * "without forceNew" mention is exempt only when it is not negated (no
 * not/never/don't/do not/must earlier in the same sentence) — "Do NOT call
 * teams_login without forceNew" still warns about the destructive flag.
 * Every remaining mention needs a whole approval phrase earlier in its own
 * sentence, within 100 chars.
 */

const MENTION = /force[_\s-]?new|--force\b/gi;
const EXEMPT = /without forceNew/gi;
const NEGATION = /\b(?:not|never|don't|do not|must)\b/i;
const APPROVAL = /ask the user to approve|with the user'?s approval|user approval only|only with the user'?s approval/i;
const SENTENCE_END = /[.!?]\s/g;
const APPROVAL_WINDOW = 100;

/** Index where the sentence containing `index` starts (just after the last boundary). */
function sentenceStart(text: string, index: number): number {
  let start = 0;
  for (const boundary of text.matchAll(SENTENCE_END)) {
    const end = (boundary.index ?? 0) + boundary[0].length;
    if (end <= index) start = end;
    else break;
  }
  return start;
}

export function recommendsForceNewWithoutApproval(text: string): boolean {
  let stripped = '';
  let last = 0;
  for (const match of text.matchAll(EXEMPT)) {
    const index = match.index ?? 0;
    const before = text.slice(sentenceStart(text, index), index);
    if (NEGATION.test(before)) continue;
    stripped += text.slice(last, index);
    last = index + match[0].length;
  }
  stripped += text.slice(last);

  for (const match of stripped.matchAll(MENTION)) {
    const index = match.index ?? 0;
    const window = stripped.slice(Math.max(sentenceStart(stripped, index), index - APPROVAL_WINDOW), index);
    if (!APPROVAL.test(window)) return true;
  }
  return false;
}

/**
 * Values of string-like literals in a TS source file, via the compiler API:
 * StringLiteral, NoSubstitutionTemplateLiteral and template head/middle/tail
 * parts. Bare identifiers ('forceNew' as a value) are skipped. node.text is
 * already unescaped.
 */
export function stringLiteralValues(fileName: string, text: string): string[] {
  const source = ts.createSourceFile(fileName, text, ts.ScriptTarget.Latest, true);
  const values: string[] = [];
  const visit = (node: ts.Node): void => {
    if (
      ts.isStringLiteral(node) ||
      ts.isNoSubstitutionTemplateLiteral(node) ||
      node.kind === ts.SyntaxKind.TemplateHead ||
      node.kind === ts.SyntaxKind.TemplateMiddle ||
      node.kind === ts.SyntaxKind.TemplateTail
    ) {
      const text = (node as ts.LiteralLikeNode).text;
      if (!/^\w+$/.test(text)) values.push(text);
    }
    ts.forEachChild(node, visit);
  };
  visit(source);
  return values;
}
