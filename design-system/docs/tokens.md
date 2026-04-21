# Токены

Все токены объявлены в `design-system/tokens.css` на `:root` и переопределяются в `html[data-theme="dark"]`.

## Цвета

| Токен | Назначение |
|-------|------------|
| `--bg0`, `--bg1`, `--bg2` | Фоны страницы и панелей |
| `--card`, `--cardBorder` | Карточки и обводка |
| `--hairline` | Тонкие разделители |
| `--text`, `--muted`, `--text-subhead` | Основной текст, вторичный, подзаголовки |
| `--primary`, `--primary2` | Акцент и мягкий акцент |
| `--inputBg`, `--inputBorder` | Поля ввода |
| `--strategy-sidebar-sep` | Разделители сайдбара стратегии (только тёмная тема) |

## Типографика

Базовая шкала (числа в `rem`):

- `--font-heading`, `--font-large`, `--font-medium`, `--font-small`

Семантические алиасы (используются в компонентах и layout):

- `--fs-widget-title`, `--fs-widget-hint`, `--fs-ui`, `--fs-body`, `--fs-small`, `--fs-xs`, `--fs-stat`, `--fs-title`, `--fs-subtitle`

Шрифт стека задаётся в `foundation.css` на `body` (системный стек Apple + fallback).

## Радиусы и тени

- `--radius`, `--radius2`
- `--shadow`, `--shadow-sm`

## Прочее

- `--ui-tap-scale` — масштаб для касаний / крупных контролов
- `--ring` — цвет фокус-кольца (если понадобится явный focus ring)
- `--app-shell-header-offset` — отступ под шапку shell

## Использование в коде

Предпочитайте токены вместо литералов:

```css
.myBlock{
  color: var(--text);
  border: 1px solid var(--cardBorder);
  font-size: var(--fs-small);
}
```
