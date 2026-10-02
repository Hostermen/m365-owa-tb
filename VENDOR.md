# Vendored third-party libraries

Bundled copies are unmodified upstream releases. Do not edit the files below;
update them only by replacing with a newer pristine upstream copy and updating
the Version/URL entries here.

options_ui/particles.min.js:
- Version: upstream master @ d01286d (post-v2.0.0, banner still reads v2.0.0)
- URL: https://raw.githubusercontent.com/VincentGarreau/particles.js/d01286d6dcd61f497d07cc62bd48e692f6508ad5/particles.min.js
- License: MIT (Vincent Garreau)

experiments/calendar/ (calendar_calendars, calendar_items, calendar_provider, calendar_timezones):
- Source: Thunderbird published WebExtension Experiment API draft "calendar"
  (https://github.com/thunderbird/webext-experiments/tree/main/calendar)
- Upstream commit: b7f7cb3e76807903a785a03784d6e7df7b213f21 (2026-04-20)
- Draft add-on version: 2.2.0
- License: MPL-2.0 (Mozilla Public License 2.0)
- Status: Unmodified verbatim copies. Do not edit these files; update only by
  replacing the whole tree with a newer pristine upstream copy and bumping the
  commit/URL entries here. The draft additionally ships two demo-UI experiments
  (calendarItemAction, calendarItemDetails) that are intentionally not included
  here: they monkeypatch core calendar dialogs on startup and are unrelated to
  this add-on's sync functionality.
