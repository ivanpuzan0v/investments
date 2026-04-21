# Компоненты

Краткий справочник по примитивам в `design-system/components/`. Разметка остаётся в `index.html` / шаблонах в `app.js`; здесь только классы и поведение.

## Таблица — `tables.css`

**Классы:** `.tableWrap`, `.table`, `.tableEmptyCell`

- Обёртка с горизонтальным скроллом, скруглением и фоном.
- Шапка: сортировка по `data-sort-field` / `data-sort-dir` (`asc` | `desc`).
- Строки: hover, последняя без нижней границы.

**Пример:**

```html
<div class="tableWrap">
  <table class="table">
    <thead>...</thead>
    <tbody>...</tbody>
  </table>
</div>
```

Специализированные таблицы (например `strategyPlansPane__table`) наследуют `.table` и добавляют свои модификаторы в `styles.css`.

## Карточка и виджет — `widget-card.css`

**Классы:** `.card`, `.widget`, `.widget__header`, `.widget__title`, `.widget__hint`, `.widget__headerRight`, `.widget__footerActions`

- `.card` — обводка, фон, тень, скругление.
- `.widget` — внутренние отступы карточки с контентом.
- Правый блок шапки с селектами: модификатор `.widget__headerRight--stackedSelects`.

## Ряд действий в строке — `row-actions.css`

**Класс:** `.buyRowActions`

Flex-ряд иконок в ячейке таблицы (копирование, календарь, удаление и т.д.).

## Икон-кнопка — `icon-button.css`

**Класс:** `.iconBtn`

Круглая кнопка под SVG или текст «✕». Не путать с `.btn` из `buttons.css`.

**Модификаторы по атрибуту:**

- `[data-action="remove"]` — красный hover.

Фокус без видимого кольца: `:focus:not(:focus-visible)` задан в этом же файле. Глобальный сброс outline для `.iconBtn:focus` остаётся в конце `styles.css`.

## Текстовые кнопки — `buttons.css` (корень репозитория)

Не входит в папку `design-system/`, но является частью единого UI:

- `.btn`, `.btn--primary`, `.btn--secondary`, `.btn--destructive` (+ алиасы ghost/danger).

См. комментарий в начале `buttons.css`.

## Что пока в `styles.css`

Сложные или единичные блоки (модалки, графики, layout стратегии, оверрайды под конкретные экраны) остаются в основном файле стилей. По мере повторения выносятся в `design-system/components/`.
