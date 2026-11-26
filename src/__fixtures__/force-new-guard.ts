import ts from 'typescript';

/**
 * Test helper: true when a piece of advice recommends forceNew without
 * requiring user approval.
 *
 * Mentions are forceNew/force_new/force-new, the CLI --force flag, or
 * force = true / force: true. A "without forceNew" mention is exempt only
 * when it is not negated (no not/never/don't/do not earlier in the same
 * sentence) — "Do NOT call teams_login without forceNew" still warns about
 * the destructive flag. Every remaining mention needs a whole approval
 * phrase earlier in its own sentence, within 100 chars; sentence boundaries
 * are . ! ? ; : and newlines, so approval in a different clause does not
 * count.
 */

const MENTION = /force[_\s-]?new|--force\b|force\s*[:=]\s*true/gi;
const EXEMPT = /without forceNew/gi;
const NEGATION = /\b(?:not|never|don't|do not)\b/i;
const APPROVAL = /ask the user to approve|with the user'?s approval|user approval only|only with the user'?s approval/i;
const SENTENCE_END = /[.!?;:]\s|\n/g;
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

/** Flattened operands of a `+` chain (left-associative nesting included). */
function plusOperands(node: ts.Expression): ts.Expression[] {
  if (ts.isBinaryExpression(node) && node.operatorToken.kind === ts.SyntaxKind.PlusToken) {
    return [...plusOperands(node.left), ...plusOperands(node.right)];
  }
  return [node];
}

/**
 * Values of string-like literals in a TS source file, via the compiler API.
 * Template expressions and `+` chains that contain a string literal are
 * emitted as ONE joined value — string-literal and no-substitution-template
 * spans inline (even \w+ ones), everything else becomes '…' — so advice
 * split across an interpolation or concatenation is still scanned as the
 * single string the runtime produces. Standalone \w+ literals ('forceNew'
 * as a property name) stay skipped. node.text is already unescaped.
 *
 * Accepted gap: same-file const identifiers are not inlined, so
 * `const FLAG = 'forceNew'; 'call teams_login with ' + FLAG` is not caught.
 */
export function stringLiteralValues(fileName: string, text: string): string[] {
  const source = ts.createSourceFile(fileName, text, ts.ScriptTarget.Latest, true);
  const values: string[] = [];
  const isStringPart = (node: ts.Node): node is ts.StringLiteral | ts.NoSubstitutionTemplateLiteral =>
    ts.isStringLiteral(node) || ts.isNoSubstitutionTemplateLiteral(node);
  const isPlusChain = (node: ts.Node): boolean =>
    ts.isBinaryExpression(node) && node.operatorToken.kind === ts.SyntaxKind.PlusToken;

  const visit = (node: ts.Node): void => {
    if (ts.isTemplateExpression(node)) {
      let joined = node.head.text;
      for (const span of node.templateSpans) {
        joined += isStringPart(span.expression) ? span.expression.text : '…';
        joined += span.literal.text;
      }
      if (!/^\w+$/.test(joined)) values.push(joined);
      for (const span of node.templateSpans) visit(span.expression);
      return;
    }

    if (isPlusChain(node)) {
      const operands = plusOperands(node as ts.BinaryExpression);
      if (operands.some(isStringPart)) {
        const joined = operands
          .map(operand => (isStringPart(operand) ? operand.text : '…'))
          .join('');
        if (!/^\w+$/.test(joined)) values.push(joined);
        // Literal operands are inlined above; still scan non-literal
        // operands for literals nested inside calls/conditionals.
        for (const operand of operands) {
          if (!isStringPart(operand) && !isPlusChain(operand)) visit(operand);
        }
        return;
      }
    }

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
