/**
 * Test helper: true when a piece of advice recommends forceNew without
 * requiring user approval.
 *
 * Mentions are forceNew/force_new/force-new or the CLI --force flag. A
 * "without forceNew" mention is exempt only when it is not negated (no
 * not/never/don't/do not/must in the ~25 chars before it) — "Do NOT call
 * teams_login without forceNew" still warns about the destructive flag.
 * A mention is allowed only alongside a whole approval phrase.
 */

const MENTION = /force[_\s-]?new|--force\b/i;
const EXEMPT = /without forceNew/gi;
const NEGATION = /\b(?:not|never|don't|do not|must)\b/i;
const APPROVAL = /ask the user to approve|with the user'?s approval|user approval only|only with the user'?s approval/i;
const EXEMPT_LOOKBEHIND = 25;

export function recommendsForceNewWithoutApproval(text: string): boolean {
  let stripped = '';
  let last = 0;
  for (const match of text.matchAll(EXEMPT)) {
    const index = match.index ?? 0;
    const before = text.slice(Math.max(0, index - EXEMPT_LOOKBEHIND), index);
    if (NEGATION.test(before)) continue;
    stripped += text.slice(last, index);
    last = index + match[0].length;
  }
  stripped += text.slice(last);

  if (!MENTION.test(stripped)) return false;
  return !APPROVAL.test(text);
}
