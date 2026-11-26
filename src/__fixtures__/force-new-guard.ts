/**
 * Test helper: true when a piece of advice recommends forceNew without
 * requiring user approval. Mentions of "without forceNew" are stripped first,
 * so only positive recommendations are flagged.
 */
export function recommendsForceNewWithoutApproval(text: string): boolean {
  const stripped = text.replace(/without forceNew/gi, '');
  return /forceNew/i.test(stripped) && !/approv/i.test(stripped);
}
