function bc_getMinimumChangeStandard_() {
  return `
MINIMUM CHANGE STANDARD — HARD LOCK:
1. Change the fewest governed items possible.
2. Within each changed item, change the fewest words or structural elements possible.
3. Preserve every unaffected word, heading, section, order and governed decision exactly.
4. Do not make stylistic, SEO, readability, variety or preference-based improvements unless explicitly required to resolve the failed check.
5. If one change resolves the failure, do not make a second change.
6. Never rewrite a complete item when adding, removing or replacing a smaller phrase or element will resolve the failure.
7. Do not alter an item merely because another version sounds better.
8. Every change must be directly traceable to a specific failed check.
`.trim();
}