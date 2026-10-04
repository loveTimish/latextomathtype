# Printable student and teacher papers

The normal JSON Word/PDF export routes support optional print layout, page numbering, conservative stem flow and exam typography. Legacy formatting remains the default.

```json
{
  "paper": {
    "name": "Example student paper",
    "compactLayout": false,
    "printLayout": true,
    "pageNumbers": true,
    "typography": "exam"
  },
  "sections": [{"questions": [{
    "serialNumber": 1,
    "questionType": 5,
    "content": "Find the total<br/>number of objects.\n\nExplain your method.",
    "contentFlow": "flow",
    "pageBreakBefore": false,
    "answerSpaceLines": 6
  }]}]
}
```

## Layout options

- `paper.printLayout`: keep a bounded opening group with the first working lines or first solution step. Later answer-space lines and long questions can paginate. Short introductions ending in a colon stay with the next paragraph. Short final solution steps stay with the answer. Individual paragraphs use `w:keepLines`; there is no unbounded keep-next chain across all working lines.
- `paper.pageNumbers`: add neutral `PAGE` / `NUMPAGES` fields in the non-compact layout. The compact footer is unchanged.
- `question.answerSpaceLines`: integer 0–12, default 3. Used for unanswered, numbered calculation/solution questions (types 5/6) under the existing answer-space conditions. Printable/exam working lines are exactly 20 pt high, without additional paragraph spacing.
- `question.pageBreakBefore`: start the question on a new page. The break belongs to the question's first paragraph, including an optional phase label or an empty stem preceding `question.images`. It is emitted once and does not carry into subsequent questions.

## Conservative stem flow

`question.contentFlow` accepts `legacy` (the default) or `flow`. Flow only joins a single source newline or single HTML `br` when it appears to be a soft prose wrap. Adjacent Chinese text is joined without an inserted space; English words keep a separating space. Blank lines, explicit block paragraphs, subquestion/option markers, short colon introductions, and standalone formula lines remain separate. Mathematical delimiters and array/aligned row separators are protected before paragraph splitting.

This setting applies only to `question.content`. It does not flatten `solution`, `analyze` or option paragraphs. Explicit solution/analysis `br` steps remain separate, including with escaped dollar signs. HTML wrapping inside a delimited formula remains whitespace inside that same formula.

Flow mode rejects HTML `img` and `table` blocks with a validation error. Supply images through the existing `question.images` field, or use the layout export route for supported structured layouts. Flow does not add a new asset-loading mechanism or silently flatten these blocks.

## Exam typography

`paper.typography` accepts `legacy` (the default) or `exam`. Exam is an opt-in print preset:

- 12 pt black body text, Times New Roman for Latin text and 宋体 for CJK
- 24 pt hanging question numbers; continuation text and standalone formulas share the body left edge
- 9.5 pt metadata on a 14 pt line; ordinary body paragraphs use a 19 pt line
- Formula and image paragraphs use at-least spacing, so tall fractions and other objects can expand the line rather than being clipped
- A4 paper with 22 mm side margins, 18 mm top/bottom margins and a 425-twip footer distance
- Sized editable formula previews use the paragraph's 12 pt body or 9.5 pt metadata size; a standalone display formula with only trailing punctuation stays in display style

`paper.previewBackground` accepts `transparent` (default) or `white`. White requires exam typography and is an explicit compatibility option for opaque preview backgrounds; it does not change the editable equation content.

Exam also enables the bounded paragraph keeps described above. Page numbering and stem flow remain independent options. Exam cannot be combined with `compactLayout` or explicit source formula metrics. Unknown mode values, incompatible options and invalid working-space counts return HTTP 400 with `INVALID_EXPORT_REQUEST`.

## Validation

Synthetic regressions cover default and opt-in behavior, field serialization, bounded working space, independent flags, builder reuse, question boundaries, page-break placement, block/word/math boundaries, font sizes, line rules and source-metric conflicts. These checks verify DOCX structure. Actual page counts and visual pagination also depend on the renderer and installed fonts; Word/MathType native GUI editing remains a separate validation boundary.
