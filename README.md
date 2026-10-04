# @mk/film-sheet

An Excel-like spreadsheet web component with a fast formula engine, built on
[Film](https://www.npmjs.com/package/@mk/film).

## Use

```sh
npm i @mk/film-sheet @mk/film lit
```

```js
import '@mk/film/css/themes/default/index.css'
import '@mk/film-sheet'
```

```html
<film-sheet style="height: 480px"></film-sheet>
```

The formula engine is also available on its own, without the DOM, from
`@mk/film-sheet/engine`. See [docs/API.md](docs/API.md).

## Develop

Requires Node 22.12 or later.

```sh
npm install
npx playwright install chromium   # browser tests
npm run dev                       # demo at http://localhost:5173
npm run check                     # typecheck, lint, unit + browser tests
```

See [docs/ARCHITECTURE.md](docs/ARCHITECTURE.md) for how it works.

## License

[MIT](LICENSE)
