# Complex Roundtrip Fixture

Marker: `E2E-COMPLEX-ROUNDTRIP-2026-08-06`

Intro paragraph with **bold**, *italic*, ~~strike~~, and `inline_code`, plus a link to [WATCHER](https://example.com/watcher) and snake_case_id `order_status_code`.

## Nested lists

- Alpha
  - Alpha-child
    - Alpha-grandchild
- Beta

1. First ordered
2. Second ordered
   1. Nested ordered
   2. Nested ordered two

### Table and quote

| Quarter | Revenue | Notes |
|---------|---------|-------|
| Q1 | 1200000 | baseline |
| Q2 | 1350000 | growth |

> Important blockquote: preserve me across docx/odt/rtf/epub.

```ts
export function hello(name: string): string {
  return `Hello, ${name}`;
}
```

---

## Slide Two For PPTX

Body for second slide after `##`.

- Bullet one on slide 2
- Bullet two on slide 2

| A | B |
|---|---|
| 1 | 2 |

Final line with marker again: `E2E-COMPLEX-ROUNDTRIP-2026-08-06`.
