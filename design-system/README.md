# Design system

Слой дизайн-системы для приложения «Планирование инвестиций»: токены, базовая типографика/фон страницы и переиспользуемые примитивы интерфейса.

## Подключение

Точка входа — `design-system/design-system.css`. Она подключается из корневого `styles.css` через `@import` в самом начале файла, чтобы все последующие правила приложения могли использовать токены и базовые компоненты.

```css
@import url("./design-system/design-system.css?v=1");
```

Кэш-бастинг версии меняйте в `index.html` у ссылки на `styles.css`.

## Состав

| Файл | Назначение |
|------|------------|
| `tokens.css` | Цвета, шрифты (шкала), радиусы, тени, переменные светлой/тёмной темы |
| `foundation.css` | `box-sizing`, базовые `html`/`body`, фон страницы |
| `components/tables.css` | `.tableWrap`, `.table`, сортировка в шапке, строки |
| `components/widget-card.css` | `.card`, `.widget`, шапка виджета, футер действий |
| `components/row-actions.css` | `.buyRowActions` |
| `components/icon-button.css` | `.iconBtn` и варианты hover для удаления |

## См. также

- **Текстовые кнопки** (primary / secondary / destructive): отдельный файл в корне проекта — `buttons.css` (подключается в `index.html` после `styles.css`).
- Документация по токенам: [docs/tokens.md](./docs/tokens.md).
- Документация по компонентам: [docs/components.md](./docs/components.md).

## Расширение

При появлении нового повторяющегося паттерна:

1. Добавьте CSS в `design-system/components/<имя>.css`.
2. Подключите файл в `design-system/design-system.css` (соблюдайте порядок: сначала то, от чего зависят остальные).
3. Удалите дублирующие правила из `styles.css`.
4. Кратко опишите компонент в `docs/components.md`.
