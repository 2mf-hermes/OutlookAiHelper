# DESIGN NOTES — macOS Glass pass (0928.1)

## Direction
- Platform language: macOS (not iOS-only). Translucent vibrancy materials.
- Dropdown: focus/keyboard + selection = inverted blue fill, white text.
- Glass: see-through white layers over Morandi canvas; light shadows.
- Group lists: inset cards, no full-card invert (that looked huge).

## Safe patterns
- Combo template via XamlReader.Parse + standard XAML Binding (RelativeSource only).
- C# FrameworkElementFactory: never mix Binding.Source + RelativeSource.
- No BlurEffect, no acrylic P/Invoke.

## Tokens
- Accent/selection: #0078D4 (Outlook-aligned)
- Inverted item: #0078D4 @ ~90% alpha, white text
- Popover: #EBF5F5F7 macOS vibrancy gray
- Menu row radius 6, popover radius 8
