/**
 * The keyboard shortcuts simple_calc handles, in one table: the Keyboard shortcuts reference
 * (Help, Ctrl+/) is generated from it. Keep it in step with the key handlers in App.tsx.
 */
export interface ShortcutEntry {
  /** Keys as shown to the user ("Ctrl+Shift+L"); alternatives separated by " or ". */
  keys: string
  label: string
}

export interface ShortcutGroup {
  title: string
  items: ShortcutEntry[]
}

export const SHORTCUT_GROUPS: ShortcutGroup[] = [
  {
    title: 'General',
    items: [
      { keys: 'Ctrl+O', label: 'Open a spreadsheet' },
      { keys: 'Ctrl+S', label: 'Save' },
      { keys: 'Ctrl+Shift+S', label: 'Save as' },
      { keys: 'Ctrl+Shift+E', label: 'Export as' },
      { keys: 'Ctrl+P', label: 'Print' },
      { keys: 'Ctrl+Z', label: 'Undo' },
      { keys: 'Ctrl+Y or Ctrl+Shift+Z', label: 'Redo' },
      { keys: 'Ctrl+F', label: 'Find in this sheet' },
      { keys: 'Ctrl+H', label: 'Find and replace' },
      { keys: 'F5 or Ctrl+G', label: 'Go to a cell, range or name' },
      { keys: 'Alt+/', label: 'Search the menus' },
      { keys: 'Ctrl+/', label: 'Keyboard shortcuts' },
    ],
  },
  {
    title: 'Move around',
    items: [
      { keys: 'Arrow keys', label: 'Move one cell' },
      { keys: 'Ctrl+Arrow', label: 'Move to the edge of the data' },
      { keys: 'Home', label: 'Start of the row' },
      { keys: 'Ctrl+Home', label: 'Cell A1' },
      { keys: 'Ctrl+End', label: 'Last used cell' },
      { keys: 'Page Up or Page Down', label: 'One screen up or down' },
      { keys: 'Alt+Page Up or Alt+Page Down', label: 'One screen left or right' },
      { keys: 'Ctrl+Page Up or Ctrl+Page Down', label: 'Previous or next sheet' },
      { keys: 'Ctrl+Backspace', label: 'Show the active cell' },
      { keys: 'Ctrl+[', label: 'Go to the cells a formula refers to' },
      { keys: 'Alt+Enter', label: 'Follow the link in the active cell' },
    ],
  },
  {
    title: 'Select',
    items: [
      { keys: 'Shift+Arrow', label: 'Extend the selection' },
      { keys: 'Ctrl+Shift+Arrow', label: 'Extend to the edge of the data' },
      { keys: 'Ctrl+A', label: 'Select the data region, then the whole sheet' },
      { keys: 'Ctrl+Space', label: 'Select whole columns' },
      { keys: 'Shift+Space', label: 'Select whole rows' },
      { keys: 'Ctrl+Shift+8', label: 'Select the current region' },
      { keys: 'Shift+Backspace', label: 'Select only the active cell' },
      { keys: 'Ctrl+.', label: 'Move to the next corner of the selection' },
    ],
  },
  {
    title: 'Enter and edit',
    items: [
      { keys: 'F2', label: 'Edit the active cell' },
      { keys: 'Enter or Shift+Enter', label: 'Finish the entry and move down or up' },
      { keys: 'Tab or Shift+Tab', label: 'Finish the entry and move right or left' },
      { keys: 'Esc', label: 'Cancel the entry' },
      { keys: 'Alt+Enter', label: 'New line in the cell (while editing)' },
      { keys: 'Ctrl+Enter', label: 'Fill the selection with the entry' },
      { keys: 'Alt+Down', label: 'Pick from a drop-down list of the column’s entries' },
      { keys: 'F4', label: 'Switch between relative and absolute references' },
      { keys: 'Alt+=', label: 'AutoSum' },
      { keys: 'Ctrl+;', label: 'Insert today’s date' },
      { keys: 'Ctrl+Shift+:', label: 'Insert the current time' },
      { keys: 'Ctrl+\'', label: 'Copy the formula from the cell above' },
      { keys: 'Ctrl+Shift+"', label: 'Copy the value from the cell above' },
      { keys: 'Ctrl+D or Ctrl+R', label: 'Fill down or right' },
      { keys: 'Delete', label: 'Clear the contents' },
      { keys: 'Ctrl+K', label: 'Insert or edit a link' },
      { keys: 'Shift+F2', label: 'Add or edit a note' },
      { keys: 'Ctrl+Alt+M', label: 'Add a comment' },
    ],
  },
  {
    title: 'Clipboard and cells',
    items: [
      { keys: 'Ctrl+C', label: 'Copy' },
      { keys: 'Ctrl+X', label: 'Cut' },
      { keys: 'Ctrl+V', label: 'Paste' },
      { keys: 'Ctrl+Shift+V', label: 'Paste values only' },
      { keys: 'Ctrl+Alt+V', label: 'Paste special' },
      { keys: 'Ctrl+Shift+=', label: 'Insert cells' },
      { keys: 'Ctrl+-', label: 'Delete cells' },
      { keys: 'Ctrl+9 or Ctrl+0', label: 'Hide rows or columns' },
      { keys: 'Ctrl+Shift+9 or Ctrl+Shift+0', label: 'Unhide rows or columns' },
      { keys: 'Shift+F11', label: 'Insert a sheet' },
    ],
  },
  {
    title: 'Format',
    items: [
      { keys: 'Ctrl+B', label: 'Bold' },
      { keys: 'Ctrl+I', label: 'Italic' },
      { keys: 'Ctrl+U', label: 'Underline' },
      { keys: 'Ctrl+5 or Alt+Shift+5', label: 'Strikethrough' },
      { keys: 'Ctrl+1', label: 'Format cells' },
      { keys: 'Ctrl+Shift+1', label: 'Number format (#,##0.00)' },
      { keys: 'Ctrl+Shift+4', label: 'Currency format' },
      { keys: 'Ctrl+Shift+5', label: 'Percent format' },
      { keys: 'Ctrl+Shift+7', label: 'Outline border' },
      { keys: 'Ctrl+Shift+-', label: 'Remove borders' },
      { keys: 'Ctrl+\\', label: 'Clear formatting' },
    ],
  },
  {
    title: 'Data and formulas',
    items: [
      { keys: 'Ctrl+Shift+L', label: 'Create or remove a filter' },
      { keys: 'Ctrl+T', label: 'Format as table' },
      { keys: 'Ctrl+F3', label: 'Named ranges' },
      { keys: 'Alt+Shift+Right or Alt+Shift+Left', label: 'Group or ungroup rows' },
      { keys: 'Alt+F5 or Ctrl+Alt+F5', label: 'Refresh the pivot table or all pivot tables' },
      { keys: 'Ctrl+`', label: 'Show formulas' },
      { keys: 'F9', label: 'Calculate now' },
      { keys: 'Shift+F9', label: 'Calculate the active sheet' },
      { keys: 'Ctrl+Alt+F9', label: 'Recalculate every formula' },
    ],
  },
  {
    title: 'View',
    items: [
      { keys: 'Ctrl+Mouse wheel', label: 'Zoom in or out' },
      { keys: 'Shift+F10 or Menu key', label: 'Open the context menu' },
    ],
  },
]

/** Entries whose keys or description contain the query (case-insensitive). */
export function filterShortcuts(groups: ShortcutGroup[], query: string): ShortcutGroup[] {
  const needle = query.trim().toLocaleLowerCase()
  if (!needle) return groups
  return groups
    .map((group) => ({
      ...group,
      items: group.items.filter((item) => item.label.toLocaleLowerCase().includes(needle) || item.keys.toLocaleLowerCase().includes(needle) || group.title.toLocaleLowerCase().includes(needle)),
    }))
    .filter((group) => group.items.length > 0)
}
