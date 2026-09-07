# window

Control Excel window visibility, position, and status bar.

## Tool

`window`

## Visibility

### Show / hide Excel

```json
{
  "action": "show",
  "session_id": "abc123"
}
```

```json
{
  "action": "hide",
  "session_id": "abc123"
}
```

### Bring to front

```json
{
  "action": "bring-to-front",
  "session_id": "abc123"
}
```

## State & Info

### Get window info

```json
{
  "action": "get-info",
  "session_id": "abc123"
}
```

Returns visibility, position, size, and foreground status.

### Set window state

```json
{
  "action": "set-state",
  "session_id": "abc123",
  "state": "maximized"
}
```

- `state`: `normal` | `minimized` | `maximized`

### Set window position

```json
{
  "action": "set-position",
  "session_id": "abc123",
  "left": 0,
  "top": 0,
  "width": 960,
  "height": 1080
}
```

### Arrange window

```json
{
  "action": "arrange",
  "session_id": "abc123",
  "preset": "right-half"
}
```

- `preset`: `left-half`, `right-half`, `top-half`, `bottom-half`, `center`, `full-screen`

## Topmost / Status Bar

### Pin / unpin window

```json
{
  "action": "set-topmost",
  "session_id": "abc123"
}
```

```json
{
  "action": "unset-topmost",
  "session_id": "abc123"
}
```

### Set status bar text

```json
{
  "action": "set-status-bar",
  "session_id": "abc123",
  "text": "Building dashboard..."
}
```

### Clear status bar

```json
{
  "action": "clear-status-bar",
  "session_id": "abc123"
}
```

## Notes

- Use `set-topmost` for side-by-side coding and debugging.
- Use `arrange` presets for split-screen layouts (e.g. Excel on one half, AI assistant on the other).
- Status bar text gives users real-time feedback during long-running operations.
