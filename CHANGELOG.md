# Changelog

All notable changes to **XLIDE: VBA for VS Code** are documented here.

## [Unreleased]

- **A container's children were already members; now they are pinned.** #57
  reported that a Frame's and a MultiPage's controls are invisible to the
  analyzer, because the two-buffer designer reader can only see the form's
  own `f` stream. Measured against a nested form: they are visible - the
  member surface is read by the package walker, which recurses into every
  child storage, and code touching `PickAir` inside a Frame or `Agree` on a
  Page analyzes clean. The flat reader is now only the fallback for a form
  the walker cannot parse, and says so. What the check DID find is that a
  container was listed twice, once with its own record and again in the
  sweep that catches a container carrying no record; completion deduped it
  downstream, so nothing showed, but the member list was wrong and is fixed.
  The fixture the old tests used has no containers at all, so nesting was
  never exercised: it is now, against an Excel-authored nested form.

## [5.0.1] - 2026-08-31

- **A multi-selection is one thing.** Ctrl-click, Shift-click, or drag a band
  across the form to pick several controls; whatever they then do, they do
  together. Ctrl+C and Ctrl+V copy the selection - a container brings its
  children, renamed and still at their own offsets - Delete removes all of
  it, and dragging or arrow-nudging any member carries the rest with their
  spacing intact, clamped as a group at the form's edge. Each is ONE write
  and ONE undo step. Copy holds NAMES rather than bytes, so pasting after
  the source is gone fails out loud instead of pasting a ghost.

- **Align and Make Same Size, on a multi-selection.** Ctrl-click or
  Shift-click adds controls to the selection; the one picked FIRST stays the
  anchor - it keeps the handles, it never moves, and it is the ruler the
  others are measured against. Align lefts, centers, rights, tops, middles
  or bottoms, and make the selection the anchor's width, height or both.
  However many controls move, it is ONE write and ONE undo step: the canvas
  computes the target geometry in points and hands the engine a single
  batch, which refuses outright if any name in it is unknown rather than
  half-applying an align.

- **Opening a form in the VBE does not reorder it.** The depth work raised
  the question of whether a VBE visit rewrites what XLIDE wrote. Measured:
  a form opened in the VBE designer, dirtied and saved comes back with its
  site order byte-identical. The reshuffle seen earlier belongs to Excel's
  own ZOrder call recomputing depth, not to opening a form.

- **Tab order and depth, the two orders geometry does not show.** The
  designer gains the VBE's Tab Order dialog - the whole surface listed in
  tab sequence, Move Up and Move Down, applied as ONE gesture so it is one
  undo step and one write - and Front / Back buttons that set a control's
  depth. Depth is real, not cosmetic: MSForms stores it as the sibling
  order in the saved form, last site on top. That was established by
  letting Excel do it (calling ZOrder on a control in the live designer and
  reading back what it wrote), and then checked in both directions - XLIDE
  writes a control to the front natively, Excel opens the form, sends a
  different one to the front, and its result extends ours under the same
  rule. Tab order and depth are independent: reordering depth leaves every
  TabIndex untouched, which is exactly what MSForms does.

## [5.0.0] - 2026-08-30

**The UserForm designer.** XLIDE edits forms now, not just the code behind
them: a canvas with the VBE's own gestures, a full properties pane, and the
form's markup - all in ONE editor over one document, with a real dirty dot,
text undo, and a save that is the only write to your workbook. The engine
underneath reads and writes MS-OFORMS designer storages natively, so none of
it needs Excel running. F5 shows the form the way the VBE does.

- **Hunt eleven: the day's new code holds.** Everything that landed today
  got its own adversarial pass and came back clean. The read/write
  classifier ran over 43,474 identifier occurrences in four real workbooks
  against an independently written oracle - zero disagreements, nothing
  thrown, 5 ms on the largest module - plus a sixty-shape matrix covering
  what the corpus oracle cannot judge: comparisons in While / Do / Until /
  Case Is that must not read as assignments, With and Me and property
  chains resolving to their terminal name, Set chains, multi-statement
  colon lines, line labels, file I/O statements, signatures, and
  bracketed FOREIGN-NAME targets. The document-highlight provider, which
  now runs on every caret move, measures 0.1 ms on the 165 KB corpus
  module. And live Excel compiled a two-form XlideRun and resolved both
  launcher subs by name, which is the first proof that the accumulating
  module is valid VBA. The shape matrix is now in the suite.

- **F5 runs what you see: the designer saves first.** The designer holds
  its gestures and markup edits as pending document changes, so pressing F5
  with a dirty form showed the LAST SAVED version in Excel - the change
  looked like it had failed. F5 now persists the form's document (and the
  focused XLIDE document) before launching, exactly as the Run-Macro
  command has always done for a dirty code module, and it does so inside
  the reopen suppression so the save cannot race the macro host. If the
  save is refused - markup mid-edit that will not parse - the launch stops
  and says so rather than quietly running the previous form.

- **F5 asks once per form, then never again.** The launcher is now one sub
  per form (XlideShow_EntryForm, XlideShow_OrderForm, ...) inside module
  XlideRun, instead of a single sub rewritten for whichever form ran last -
  so running a second form ADDS its launcher beside the first rather than
  taking the first one away, and any hand edits to the module survive. Once
  a form has its sub, F5 stops asking: there is nothing new to put in the
  workbook, so it just runs it - and skips the write entirely, meaning a
  repeat F5 no longer touches the file at all. An explicit Never still
  means never.

- **F5 on a form opens ONE Excel.** The launch writes the launcher module
  and then runs it, and the write's own Excel coordination reopens the
  workbook after a save or a close - so the workbook was opened twice for
  one keypress, once by the save's reopen and once by the macro host. The
  Run-Macro command has always held XLIDE's reopen suppressed across both
  halves for exactly this reason; the form launch now does the same, and
  additionally adopts that command's locked-workbook handling (close and
  retry under the coordination policy instead of asking you to close Excel
  by hand) and its tracking of the workbook the macro host reopened.

- **The markup pane has its colors back, and F5 asks once.** Folding the
  markup into the designer traded a real editor for a bare textarea, and
  the syntax coloring went with it. The pane now paints the dialect again -
  tags, attribute names, quoted values, punctuation, comments - on a layer
  under a transparent-text textarea, so the caret, selection, IME, and
  native undo all stay exactly as they were while the colors sit under the
  glyphs (both layers share one font, padding, and line-height rule;
  measured aligned to the pixel, with a very large document falling back to
  plain text so a giant paste cannot make typing crawl). And F5 in the
  designer raised its consent dialog TWICE: the canvas posted the launch
  and the workbench keybinding ran the same command. The keybinding is the
  single source now, with a one-at-a-time guard on the command so no second
  press can ever stack another modal.

- **Every reference now says whether it reads or writes.** Issue #55's ask:
  find-all-references answered positions but not what each one DOES, and
  the vbide's Extract Method needs exactly that. Each reference span now
  carries a kind - read, write, or readwrite - decided by the logical
  statement it sits in: the assignment family writes its target's terminal
  name (x =, x(i) =, a.b.c =, With's .field =, Set, Let, LSet, RSet, a
  Function's return assignment), declarations write the names they
  introduce (Dim, Const, parameters, the procedure name), For and For Each
  write their loop variables, ReDim writes and ReDim Preserve read-writes,
  Erase writes, Mid(s, ...) = read-writes its target, and Line Input /
  Input # / Get # write the variables they fill. Inline If classifies the
  statements after Then and Else on their own. The one gray zone stays
  honest: a variable passed where a ByRef parameter might write it remains
  a read, because claiming otherwise would need call-site signature
  resolution and be wrong for every ByVal. VS Code gets the first consumer
  in the same change: read/write occurrence shading when the caret sits on
  an identifier, writes shaded as writes.

- **Hunt ten found nothing to fix - and that is the finding.** The two
  territories no hunt had visited both came back sound. The VBA analyzer
  took a pathological sweep - 3000-deep parentheses, 2000-deep If nests, a
  one-megabyte line, ten thousand line continuations, keyword soup, null
  bytes, three hundred seeded mutants, and a scaling ladder - and stayed
  bounded everywhere (worst case 466ms, near-linear growth, the corpus's
  giant module analyzed in 119ms). And the TabStrip's tab structure,
  written through the document (a caption rename plus an added tab, five
  parallel arrays and the flags tail behind it), loaded in live Excel
  exactly as authored: three tabs, Overview / Tab2 / Extra. After ten
  hunts, every layer of the designer and the paths around it has had a
  dedicated adversarial pass.

- **A renamed page keeps its contents through the save.** Hunt nine went
  after the one identity the reconciliation had excluded: MultiPage pages.
  A page rename still saved as remove-plus-add, so the renamed page came
  back rebuilt from text alone - a picture on it died while the reprint
  stayed equal, the oracle-blind loss shape again. Pages pair FIRST of all
  now (before control moves, so a control dragged onto a renamed page
  resolves against its new name): the unambiguous single rename at the
  same position with the SAME caption - the caption is the discriminator,
  since a rename keeps it and a genuinely fresh page brings its own -
  renames the site in place; anything else falls back honestly. Live Excel
  read the save output back: two pages, the renamed one keyed by its new
  name with both children and the picture alive. Also learned from the
  bytes: the tab-strip TabNames array is internal bookkeeping (Tab3/Tab4
  in the fixture), not page identity, so a site rename needs no array
  sync. The chain fuzz now rolls page renames and captions too - 40 seeds
  green.

- **Exported forms import into the VBE now - containers, classes, pictures,
  the lot.** Hunt eight pointed the live oracle at the one output no probe
  had ever consumed: our .frm/.frx export pair. The VBE refused it -
  "Property OleObjectBlob could not be set" - and the diff against the
  VBE's own export of the same form told the whole story: a real .frx
  embeds the form's ENTIRE designer storage tree (a Frame's and a
  MultiPage's container storages, each tagged with its class CLSID, plus a
  Forms.Form CLSID on the root that OLE binds by), while ours packed only
  the flat top-level streams - containers were silently absent, and the
  in-workbook CompObj variant it copied zeroes the very CLSID the sidecar
  needs. Export now deep-copies the whole classed tree, composes the
  sidecar's own CompObj, and stamps the root CLSID and the real outer
  dimensions; import plants the tree back the same way (it too used to
  drop every container's contents). The VBE now imports our export whole:
  17 controls, frame children in place, both pictures alive. Also this
  round: a 40-seed random-chain fuzz over structural op sequences caught
  markup additions dropping their site-level say (a fresh control lost its
  Tag and had its document-spelled TabIndex reassigned) and sibling order
  diverging when a move and an add interleave - additions apply the full
  element now, and each container re-syncs its sibling order to the
  document after the diff.

- **A created control must carry a name VBA can wire.** Hunt seven noticed
  the fuzz had been ACCEPTING mutated names: markup could add a control
  named "Bad Name!" or "2Start" - names no event handler can ever exist
  for, wired into the workbook where the VBE itself would refuse them. A
  creation (an addition or a rename, from the markup pane or the
  properties pane) now demands what VBA demands: a letter first - any
  script, Japanese and accented names are legal VBA and stay legal here -
  then letters, digits, or underscores; a control an existing workbook
  already carries is matched as-is. Also verified this round: the vbide
  reads our new form attributes safely by construction (its apply resolves
  names against the live VBE dispatch, which knows ShowModal and friends
  natively), and the canvas math is exact under a form's own stored Zoom
  and under grid snap - drags, placements, and snapping all measured to
  the point in the browser harness.

- **A pasted megabyte caption can no longer corrupt the workbook.** Hunt six
  fuzzed the markup pane's new freedom - 2,015 adversarial documents (tag
  soup, truncations, null bytes, 2000-deep nesting, megabyte attributes) -
  against three contracts: throw cleanly or apply, never touch the file on
  a refusal, and stay fast. Two finds. The bad one: a 1MB caption was
  ACCEPTED, silently wrapped the format's 16-bit record length, and wrote a
  corrupt workbook the engine then refused to re-read; both the record and
  site writers now refuse oversized text with a clear message before a
  single byte lands, from the markup pane and the properties pane alike
  (large-but-legal text still works). The slow one: a 4000-control paste
  cost 4.7 seconds - the identity reconciliation rebuilt its whole index
  per document line; it maintains the index incrementally now, and the same
  paste costs 0.34s with near-linear scaling. The full fuzz closes at 2,015
  cases, zero failures, worst case 500ms.

- **The whole authored surface survives Excel's own resave - verified.**
  Hunt five put the harshest oracle available over the engine: a workbook
  carrying everything the designer can author (every form and control
  property family, the reconciled renames and reparents, fresh-added
  controls, pictures) was opened in the VBE's own DESIGNER, dirtied
  through it, and re-saved by Excel - then re-read here. Nothing was
  lost: every property, both pictures, every identity survived Excel's
  re-serialization, with three understood normalizations on Excel's side
  (font sizes snap to the raster grid, 10 becomes 9.75, exactly as the
  VBE's grid does it; a value stored at its kind default may drop to
  implicit; sibling order can shuffle under VBE edits - which the
  name-keyed diff is immune to by design). For future probes: setting a
  design property via VBComponents.Properties fails under automation
  with a focus error on ANY workbook, Excel-authored ones included -
  dirty through the designer surface (DesignerWindow plus a temp control
  add/remove) instead.

- **A renamed or reparented control no longer loses its picture on save.**
  The markup diff is keyed by name, so a rename or a move read as
  remove-plus-add - and the rebuilt control kept only what the dialect can
  spell. Its picture, its mouse icon, an ActiveX payload: silently gone
  from the SAVED workbook while the canvas still showed them (the fuzz
  oracle was provably blind to the reparent case - reprints matched while
  the bytes lost the image). A reconciliation pass now runs before the
  diff, in document order: the same name under a new container is executed
  as a real in-place MOVE, a document-only name that matches exactly one
  missing control - by kind and printed geometry in the same container, or
  by kind, size, and caption anywhere - is executed as a real RENAME, and
  anything ambiguous honestly falls back to remove-plus-add (a genuine
  delete-plus-add never inherits a dead control's picture). Frame renames
  keep their children; a child moved out of a deleted frame survives it.
  Live Excel vouches for the save output: the renamed image and the moved
  button both load pictured, inside the right parent.

- **Drag-to-reparent and form resizing actually work again - a browser
  harness now drives the real canvas.** Hunt three shimmed the webview API
  into a rendered page and replayed every gesture with synthetic pointers,
  asserting each posted message against hand-computed points. Two kills:
  the drop hit test ran AFTER the carried control lost its
  pointer-transparency, so it caught its own drop and every cross-container
  drag silently became a same-container move (the frame even highlighted,
  then nothing) - the hit test now runs first; and a local client shadow
  put the form-resize commit in its temporal dead zone, so every form
  resize THREW instead of posting - the shadow is gone and a suite pin
  keeps both orderings. Everything else measured exact: move, resize from
  every edge, nudges, click-to-place and drag-to-place, zoom-scaled drags
  (points stay true at 150%), the markup draft lifecycle, the Ctrl+Z/Y/S
  and F5 keys, and both popovers.

- **A fuzz oracle now guards the document model - and its first sweep found
  a real one.** Every property the pane offers, on every control and the
  form, is machine-checked against the two invariants the one-document
  designer stands on: the printed document reproduces the state on a fresh
  baseline, and applying a state's own print to itself changes nothing.
  The sweep (492 writes, plus revert-to-default and same-field batch
  passes) caught the empty-string asymmetry: a caption or value cleared to
  empty printed as Caption="", which no apply could re-create - and a
  frame's cleared legend quietly resurrected on save. Empty strings are
  unspoken in the dialect now, and a whole-document apply reads absence as
  empty, so cleared text stays cleared through save and undo. A trimmed
  oracle runs in the suite permanently.

- **Hunt findings in the one-unit designer, fixed the same day.** Typing in
  the markup pane dirtied the document but never repainted the canvas: the
  edit's own echo suppression swallowed the very change event that applies
  it. Gestures keep suppressing their echo; markup edits no longer do.
  Keystrokes typed in the race between a flush and the re-render could die
  with the replaced textarea; an unflushed draft now rides the webview state
  and is restored - caret, scroll, and debounce re-armed - when the new page
  arrives. Opening a clean designer no longer runs a whole parse-and-diff to
  prove the document says what the workbook says, and a typing debounce that
  fires right after a gesture no longer repeats the identical render. The
  engine loop was measured on the corpus workbooks while hunting: a gesture
  costs 11-26 ms and an undo/typing rebuild 5-6 ms, fixture through the
  2.25 MB giant-module profile - the model needs no engine-side tuning.

- **The designer and its markup are ONE editor now.** Opening a form lands
  in a single tab: the canvas and properties pane above, the markup text
  below the vbide grip - no second tab strip, no editor-group acrobatics.
  The tab carries a real dirty dot: gestures and markup edits are pending
  document changes, Ctrl+Z is ordinary text undo (from the canvas too), and
  Ctrl+S is the ONLY write to your workbook. Under the hood the canvas
  renders from a scratch copy holding the document's state applied, so what
  you see is always real bytes without your file being touched; the grip's
  arrows and drag work as before, now purely inside the editor, and typing
  in the markup pane repaints the canvas after a pause, with parse errors on
  a strip instead of a broken canvas.

- **The document now carries the WHOLE form.** The markup dialect gains the
  form-level surface the pane could already edit - BorderStyle, ScrollBars,
  Cycle, Zoom, MousePointer, the form's StdFont, and the VBFrame trio
  (ShowModal, StartUpPosition, WhatsThisButton) - printed quiet at their
  defaults and applied back, with the VBFrame twips echo kept in step on
  form resizes. And within the dialect's vocabulary the document is total
  on apply: a property edited back to its default (a re-enabled control, an
  unbolded font, a cleared ControlSource) stays that way through save and
  undo instead of silently reverting. A live-Excel probe vouches for the
  whole save path: Zoom 150, MousePointer 11, Segoe UI bold, StartUpPosition
  2, and both defaults-restorations read back exactly.

- **The form's font styles truly paint, and the picker shows real fonts.**
  Form-level Font.Bold and Font.Italic were written to the bytes but never
  drawn; the client area now wears the StdFont's weight, slant, underline,
  and strikethrough - Underline and Strikethrough join the form pane as
  editable rows, and any single style edit carries the other three instead
  of erasing them. The bytes follow the spec while they are at it: StdFont's
  fBold flag MUST be zero, so bold rides the weight (700), as Excel writes
  it. Font.Name trades its datalist - which filtered to whatever was already
  typed, showing only Tahoma - for a proper dropdown of the Windows faces
  the renderer can resolve, each previewed in its own face, with the current
  one highlighted; typing any other name still works.

- **F5 can truly RUN the form.** Excel can only run a macro, so with your
  consent XLIDE injects a small launcher module (XlideRun, safe to delete)
  and runs it - the form opens immediately, the way F5 in the VBE does. A
  modal asks the first time: Yes runs it once, Always remembers the choice,
  No just opens the workbook. The memory lives in the new
  xlide.formRun.injectShowMacro setting (ask / always / never), on the
  settings page beside the other Excel integration choices.

- **The property surface is complete, form and controls both.** The form
  pane gains BorderStyle, ScrollBars, Cycle, Zoom, MousePointer, its Font
  (Name, Size, Bold, Italic through the StdFont composer), and the VBFrame
  trio - ShowModal, StartUpPosition, WhatsThisButton - replacing the line or
  inserting it before End, spelled 0 and -1 as VB does. Controls gain
  ControlSource and RowSource, HelpContextID, MousePointer, and Alignment
  (the caption side of a CheckBox or OptionButton). One live-Excel probe
  answers for all of it: Zoom 150, ScrollBars 3, Cycle 2, StartUpPosition 2,
  Segoe UI 10, MousePointer 11, a bound ControlSource, Alignment left.

- **A form-font write could crash Excel; it cannot now.** The StdFont
  composer - never before exercised, since no fixture form stores a font -
  wrote the FontGUID's Data2 and Data3 big-endian where MS-DTYP stores them
  little-endian, and Excel answered the malformed component GUID by dying at
  form load. Found by clean-room bisection down to the single Font write;
  the GUID bytes are pinned in a test, and the fresh Font marker seeds its
  mandated 0xFFFF.

- **The Properties pane speaks each property's language** - the xlide vbide
  shapes: True/False dropdowns for the booleans, named dropdowns for the
  enums (SpecialEffect, BorderStyle, MultiSelect, ListStyle, Style,
  PictureSizeMode, PictureAlignment, Orientation, ScrollBars, TextAlign), a
  color popover with vbide's palette ramp plus the SYSTEM color names (the
  OS dialog only speaks #rrggbb, and half of a form's colors are names), a
  font-face list that stays free-text, and a numeric size field.

- **The split grip earned its keep the hard way.** Its drag is layout-tree
  pixel math now (vscode.get/setEditorLayout on this group and its column
  sibling): continuous, direction-true, and unable to move anything outside
  the designer+markup column. Applies coalesce to the latest delta - the
  earlier command-based mechanisms maximized across the whole window,
  stepped coarsely, could grab the wrong group (the direction inversion),
  flickered, and stuttered. Pointer capture keeps a drag alive across the
  iframe edge, and pointer loss (Alt+Tab) always ends it. The collapse
  arrows push the same math to its clamp.

- **The canvas renders like the native form, in a webview.** VS Code
  injects default styles into every webview; un-colored canvas text was
  inheriting the workbench foreground (gray in a dark theme), and injected
  img rules fought the caption-picture layout - both armored now. A control
  whose BackStyle is opaque wears the default ButtonFace when it stores no
  BackColor, which is what MSForms paints and what keeps the grid dots OUT
  of controls; a Frame's own BackColor finally lands on its surface.
  Buttons carry a soft raised face, edits a sunken whisper, frames the
  etched line. The Properties pane floats on the designer's gray as a
  carded panel with a resize sash and a collapse button.

- **The designer sits under its form in the tree.** Expanding a userform
  module shows a Designer row first - above the procedures, the xlide vbide
  arrangement - and clicking it opens the canvas. The markup projection also
  lost its banner comment: the document now starts at `<Form>`.

- **Designer sessions undo.** Ctrl+Z / Ctrl+Y (and Ctrl+Shift+Z) while the
  designer is focused walk a per-form history of byte-true designer
  snapshots - up to 50 gestures - restoring exactly what a gesture changed,
  structure included; the markup face follows every step. The engine grew
  the snapshot/restore pair, pinned by tests that put a moved control and
  an added control back to the exact prior bytes.

- **The designer restores after a window reload.** A webview serializer
  brings the canvas back onto its form (the markup below always restored
  natively; now both do).

- **F5 launches the host.** From the designer or the markup document, F5
  opens the workbook in its application - Excel for Excel containers, the
  owning app otherwise.

- **Form resize handles actually drag now.** The east/south/southeast
  handles seeded from an inline style the client never had, so every frame
  was a silent no-op; the seed measures the box, and the commit posts the
  tracked size (a no-move click posts nothing). Verified by driving the
  handles in the rendered page.

- **Caption pictures render the way MSForms draws them** - the rules xlide
  vbide measured off the running form: the top-left pixel is a color key
  (exact matches go transparent, the anti-aliased halo stays), an oversized
  picture stretches over the whole face with the caption underneath, one
  that fits keeps its natural size beside the caption where
  fmPicturePosition says, and position 12 stays a background. Surface
  pictures (Image, the form) still draw solid and letterbox. The titlebar
  also matches the native form now: flat white, regular Segoe UI.

- **Designer shell ergonomics.** The toolbox highlights on rollover instead
  of latching blue on click. The splitter strip at the bottom edge carries
  vbide's chrome: collapse arrows either side of the grip dots, the dots
  drag the split with the resize cursor they promise, double-click
  collapses or restores the markup below. The Properties pane gained a
  left-edge sash (width persists), a collapse button in its header, and a
  guard so clicking inside it never clears the canvas selection. Zoom: a
  preset picker plus Ctrl+wheel on a smooth exponential curve (25%-400%),
  with every pointer-to-points conversion dividing the factor back out.

- **The canvas paints real pictures.** A stored picture - on an Image
  control, on a button or label, or as the form's own background - renders
  from its actual bytes when the browser can decode the format (BMP, PNG,
  JPEG, GIF, ICO), honoring PictureSizeMode (clip, stretch, zoom) and
  PictureAlignment on Image controls. WMF and EMF have no browser decoder,
  so those keep the honest hatched placeholder rather than a guess.
  Measured in the rendered page: the fixture's 196KB BMP decodes at its
  real 256x256, zoom maps to contain, and the button wears its picture
  under its caption.

- **TextAlign, the list behaviors, ComboBox Style, and the last two font
  effects joined the dialect and the pane.** TextAlign speaks words - Left,
  Center, Right - because the stored PARAFORMAT_Alignment and VBA's
  TextAlign enum number the same three positions differently, and a bare
  number would mean two things; it applies to the kinds whose VBA surface
  has the property (a CommandButton stores an alignment but exposes none, so
  the dialect stays silent there). MultiSelect, ColumnCount, and ListStyle
  print for list and combo boxes as stored numerics; ComboBox Style="2" is
  the drop-down list (DisplayStyle 7), quiet when it is the plain combo; and
  Font.Underline / Font.Strikethrough round out the effects. Verified in
  live Excel: a centered label, a right-aligned underlined text box, a
  multi-select two-column list, and a drop-list combo all answer with the
  exact VBA enum values. The canvas draws what these set: alignment,
  underline and strikethrough, and the GrayText a disabled control wears
  over any custom color, the way the VBE paints it.

- **The everyday booleans joined the dialect: Enabled, Locked, MultiLine,
  WordWrap, AutoSize, Visible, TabStop, Default, and Cancel.** They live in
  two bitfields - the control record's VariousPropertyBits and the site's
  BitFlags - which most Office-authored controls do not store at all, so
  every write composes from the kind's file-format default straight out of
  [MS-OFORMS] 2.5.96 and 2.5.4 (measured: only text and combo boxes carry
  the record field; buttons, labels, and checks sit on defaults). The markup
  prints a flag only when it differs from that default, so quiet forms stay
  quiet and every pinned fixture projection is unchanged; the Properties
  pane always shows the effective True/False. Kind rules enforced: MultiLine
  is a TextBox property, Default and Cancel belong to CommandButton, a Label
  has no TabStop, and a Page's visibility bit is the current-page marker the
  dialect never spells. Verified in live Excel: a disabled default button, a
  locked multiline text box, an unwrapped label, and a hidden checkbox all
  load and answer with the written values.

- **Pages reorder through the markup.** List the `<Page>` elements in the
  order you want and the apply permutes everything a page's position touches
  as one move: the page sites, each page's position-tracking TabIndex, the
  tab strip's per-tab arrays with their flags, the selected-tab index
  (remapped so the same PAGE stays current), and the x bookkeeping's rows
  and PageIDs - so the page a tab names, the storage it binds, and the
  caption it wears travel together. Reorder composes with recaptions and
  additions in the same apply. This closes the last "not supported yet" in
  the page grammar; verified in live Excel, where the swapped form loads
  with both pages in the new order, their own captions, and their contents
  still bound.

- **Double-click opens the event handler.** The VBE's signature gesture:
  double-click a control on the canvas and the code face opens at its
  DEFAULT event handler - Click for buttons and labels, Change for inputs -
  creating the `Private Sub Name_Event()` stub at the end of the module when
  none exists, cursor inside. The empty face is the form itself
  (`UserForm_Click`), and a page's empty area belongs to its MultiPage. The
  stub is a document edit you save like any typed code, not a silent
  workbook write.

- **A `'''` doc block now attaches to its member through xlide's own
  directives** (#53). The doc scan was the strict one of the three comment
  grammars: any line between the block and its declaration silently detached
  the docs - including a `' @xlide-analysis-*` suppression or `' @xlide-test*`
  marker, the product's own two grammars, whose scanners were already
  tolerant in every stacking order. Directive lines are now transparent to
  the backward scan, so every order works and hover keeps the summary; a
  blank line or an ordinary comment still ends the block, as documented. A
  top-of-module block a directive separates from the first declaration now
  belongs to that declaration, never double-claimed as module docs.

- **The designer has a Properties pane.** A panel on the canvas's right edge
  follows the selection the way the VBE's Properties window does - click a
  control for its rows, the empty face (or nothing) for the form's own. The
  rows are the markup dialect's vocabulary, drawn from the same record and
  site fields the document prints - one answer behind both faces - with a
  blank meaning a property at its default. Editing a row commits ONE
  property write on Enter or blur; Escape reverts the row and keeps the
  selection; typing in the pane never nudges or deletes the selected
  control. Renaming is a row like any other: the pane follows the control to
  its new name (code that mentioned the old name is yours to update, as in
  the VBE). Verified in live Excel: a renamed button, a recolored text box,
  a recaptioned page, and the form's own caption all load and answer with
  their new values.

- **The markup document opens beneath the designer** - the xlide vbide
  arrangement, canvas above and document below - whenever Open Designer runs
  and no editor group already shows it. The two faces still track each other
  both ways.

- **Controls drag between containers.** Carry a control over a different
  surface - the form face, a Frame, a Page - and the prospective home
  outlines; the drop moves it there, the VBE's own gesture. The site, its
  record bytes, and (for a Frame or MultiPage) its entire nested storage move
  intact, so every property and child survives; position maps into the new
  surface, tab order joins the end of the target's, and group membership
  stays behind with the old container. Pages keep their MultiPage, a
  container never lands inside itself, and an ActiveX control - whose class
  entry lives in its container - stays put. Verified in live Excel: a form
  with a button moved into a frame, and that frame (button aboard) moved
  onto a page, instantiates with every parent correct and the control count
  identical to the untouched fixture's.

- **The form itself is selectable and resizable.** Click the empty face to
  activate the form: a dashed outline, and resize handles on the east, south,
  and southeast edges - the ones a form can grow by, since position is not a
  form property. Dragging resizes live and commits one write through the same
  size primitive the markup uses, so the VBFrame's twips client box stays in
  step. Escape clears any selection.

- **The snap toggles behave.** Grid and neighbor snapping are exclusive now -
  checking either clears the other, both may be off, and grid is the default;
  the two pulled the same drag toward different lines when combined. The
  choices also survive gestures: every canvas edit re-renders the page, and
  the toggles used to reset to their defaults with it. And while grid snap is
  on, every design surface paints the 6pt lattice as dots, the VBE's dotted
  face, each dot exactly where a snapped edge lands.

- **The markup document is no longer analyzed as VBA** - opening a `.form`
  painted every element as "statement outside a procedure" and offered VBA
  completion inside XML. Two gates were at fault: the VBA-document check
  claimed everything on the xlide scheme, and the language-provider selector
  carried a bare scheme match. Both now serve only the code face, and `.form`
  documents get XML colorization by extension.

- **The toolbox drags.** Press a kind and carry it: a ghost at the control's
  real default size follows the pointer onto whichever surface is underneath
  - the form, a Frame, a Page - at the grid-snapped point the drop will use,
  so what you see is what you get. A plain click still arms click-to-place.

- **"Preview Form" is now "Open Designer"**, listed above Open Form Markup -
  the canvas edits, so the old name undersold it.

- **The canvas carries the vbide designer's hover ergonomics.** The control a
  click would select outlines on hover, and only the DEEPEST hovered one does
  - a Frame never lights up under its own children. Cursors follow the rule
  the vbide settled: the hand across the whole face because every inch
  responds to a press, the four-way move only on the control a press would
  actually pick up, resize cursors only on handles, and a drag in flight
  paints the canvas with its gesture's cursor. A control being carried lifts
  above its siblings, slightly transparent, and invisible to the pointer.

- **The canvas paints with the real Windows palette, not CSS keywords.**
  Frame, page-group, and tab-group borders had vanished: modern Chromium
  computes the deprecated `ButtonShadow` CSS keyword as the same near-white
  as `ButtonFace`, so a keyword-colored border on a keyword-colored surface
  is invisible by definition - measured in the rendered page's computed
  styles. Every structural color, and every OLE system color a control
  carries, now resolves to the explicit Windows default palette value MSForms
  actually paints with.

- **The form designer has a canvas.** Right-click a UserForm for **Preview
  Form**: the designer model renders in a webview - real bounds, captions,
  colors and fonts from the binary, Frames and MultiPage pages nested, honest
  hatched placeholders where fidelity runs out (pictures, foreign controls).
  And it is not only a picture: the canvas carries the VBE designer's
  ergonomics - drag to move, eight resize handles, a toolbox that places new
  controls where you click, a 6pt grid, snapping to neighbor edges and
  centers with alignment guides, arrow-key nudges (Shift for grid steps), and
  Delete. Every gesture is one write through the same primitives the markup
  diff uses, on the authoring path live Excel verified; the canvas re-renders
  from the workbook after each, so it always shows what the bytes say. The
  markup document and the canvas track each other in both directions.

  Apply verification now covers all three hosts: a Word form and a PowerPoint
  form mutated by the engine each compiled and instantiated in their live
  application (`ok:native word edit`, `ok:native ppt edit`), closing the
  Excel-only caveat.

- **Every control on a form is a member of the form, however deeply it
  nests.** The designer-declared member surface used to carry only top-level
  controls, so `Me.PickAir` - a control inside a Frame - was reported
  undeclared in its own code-behind. The surface now walks the whole
  container tree: controls in Frames, on Pages, and the containers themselves.

- **MultiPage Pages and TabStrip tabs are editable through markup** - the one
  structural refusal worth lifting is lifted. A `<Page>` present only in the
  document is created (its nested storage, its site, its tab entry, and the
  `x` bookkeeping move together), one present only in the designer is removed,
  and page captions, tab captions, and every control ON a page edit as before.
  Tabs have no names, so their diff is positional: append, truncate, recaption.
  Reordering surviving pages stays refused by name until it is proven.

  Verified in live Excel end to end: a form whose MultiPage gained a page
  carrying a TextBox, lost another page, and whose TabStrip gained a third tab
  answered `Pages.Count = 3`, the new page's caption, its control's name, and
  the new tab's caption under `Application.Run` - alongside re-verification of
  every previously proven flow, combined in one workbook.

  Getting there surfaced four authorship rules the spec alone does not teach,
  each found by live Excel refusing a form and pinned as a test:
  control IDs are allocated from ONE counter across the whole form tree (the
  fixture's ID gaps at the root are exactly its nested controls); a fresh
  MorphData needs its reserved mask bit plus the `VariousPropertyBits` and
  populated `TextProps` Excel always writes; a fresh form must carry the empty
  class-table count word, without which its FIRST control makes fm20 parse
  garbage as class info (an empty form survives by coincidence - both misread
  values are zero); and a container's `ShapeCookie` is set only where none
  exists, never bumped, because rewriting an existing one broke a form that
  loaded fine before. A container's CompObj also names its kind: a Page
  authored with its parent MultiPage's CompObj loads and silently drops out
  of `Pages`.

- **The UserForm designer is native.** XLIDE now reads and WRITES a form's
  [MS-OFORMS] designer storage itself - no Office application in the loop.
  Right-click a workbook for **Add UserForm** and a form module for **Open
  Form Markup**: the form opens as an editable text projection in the dialect
  xlide_vbide defined (XAML-shaped elements, quoted values, `Font.Size`
  dotting, colors as `#rrggbb` or system names like `ButtonFace`), and saving
  the document applies it back as a name-keyed diff. A control present only in
  the document is added, one present only in the designer is removed, and an
  unspoken property is never touched. A parse error applies nothing.

  The engine's contract is BYTE IDENTITY, pinned across four Office-authored
  fixtures (Excel, Word, PowerPoint, and a 19-control form with fonts,
  multi-hundred-KB pictures, a Frame, and a MultiPage with nested Pages):
  parse then serialize reproduces every stream of every designer storage
  exactly, undefined padding included. Editing writes only what changed.

  New forms are authored from nothing - module, `BaseClass` registration,
  attribute header, VBFrame, minimal FormControl, Forms 2.0 CompObj - and the
  result flows straight into the existing pipeline: the module classifies as
  `userform`, carries `predeclaredId`, and its designer-declared controls
  reach completion and diagnostics.

  Verified against live Excel (verification only, per the no-COM-in-engine
  rule): a workbook whose form this engine recaptioned, moved a button in,
  and added a Label to - and a second form authored entirely from scratch -
  compiled and instantiated under `Application.Run`. That verification also
  caught a real defect the fixtures alone would not have: Excel's own
  NextAvailableID can EQUAL the highest live control ID, so a fresh ID is now
  chosen above both.

  Not in this slice, each refused with a named error rather than silently:
  creating ActiveX controls (needs class-table authoring), adding or removing
  MultiPage Pages and TabStrip tabs (captions and contents edit fine), and a
  visual canvas renderer - the markup document is the design surface, as the
  vbide plan intended.

## [4.1.7] - 2026-08-27

- **The suppression quick fix writes a directive that can reach the rule**
  (#52). It offered `disable-next-line` whatever the rule's scope, and for a
  module-scoped rule that directive cannot work: `option-explicit-missing`
  anchors its finding at the module top, so inserting a line directive above
  line 1 re-anchored the finding onto the directive itself, which
  `disable-next-line` never covers - it covers the line below. The comment
  suppressed nothing and applying the fix again stacked another one.

  The action is now chosen from the rule's own `suppressionScopes`. A rule with
  no positional scope offers `Suppress '<code>' in this file`, writing
  `disable-file` after the `Attribute` header where the directive is legal.
  Everything else keeps the next-line action unchanged.

## [4.1.6] - 2026-08-27

- **A named argument no longer breaks the block an `If` opens** (#51). A `:`
  separates statements in VBA and the scanner split on every one of them, so
  `If Range(Cell1:="a1") Is Nothing Then` was torn into `If Range(Cell1` and
  `="a1") Is Nothing Then`. Neither opens a block, so the matching `End If` was
  reported as unmatched. The colon in `:=` is the named-argument operator and is
  never a separator.

  The same root cause reached a second case the report did not name: the colons
  inside a **date literal**. `If Now > #12:30:00 PM# Then` produced the
  identical error, and `t = #12:30:00 PM#` was being split into three
  statements. The literal match is deliberately narrow - a leading digit and
  date punctuation only - so `Print #1, "a": Close #1`, where `#` opens a file
  number rather than a literal, keeps its separator.

## [4.1.5] - 2026-08-23

- **The host contract is documented and tested** (#50).
  `docs/host_contract.md` states which facts a host must supply, what each one
  costs when absent, and which of them have three states rather than two.
  `tests/hostContract.test.ts` asserts every claim in it, so the document cannot
  quietly stop being true.

- **The default-instance field is `predeclaredId`** (#50), renamed from
  `predeclared` before anything depends on it, so it mirrors the attribute a
  host reads (`Attribute VB_PredeclaredId`) rather than paraphrasing it. The
  accessor is `ProjectIndex.modulePredeclaredId`.

- **`ModuleInput` is exported** (#50). `ProjectIndex` was public but the shape
  its `setModule` accepts was not, so a host outside this repo could call the
  method without being able to import the contract it writes against - including
  `predeclared` and `implicitMembers`, the two facts a VBE `CodeModule`'s text
  cannot carry.


## [4.1.4] - 2026-08-23

- **`rs!CustomerName` is member access, not an undeclared variable.** `!` is
  VBA's default-member accessor, so a name after it belongs to the receiver
  exactly as a name after `.` does. The scanner skipped one and not the other,
  so every field read through an ADO or DAO recordset, and every Access form
  reference, was reported as undefined. The Single type suffix (`Dim x!`,
  `a! + 2`) is unaffected: a suffix is only ever followed by an operator or the
  end of the statement, never by a name.

- **`[A1]` is an Excel lookup, not an undeclared variable.** The square brackets
  are shorthand for `Application.Evaluate`, so `[A1]`, `[TaxRate]` and
  `[A1].Value` compile with nothing declared. Word has no such feature -
  measured in the VBE, `v = [foo]` with nothing declaring `foo` is a compile
  error there - so Word and PowerPoint keep the report. An absent host model is
  Excel's by default (#28) and a model that knows nothing asserts nothing, so
  both of those stay silent rather than guess.

  Across the workbook corpus this removes 10 findings and adds none, all of them
  the bracketed reserved names the torture test annotates as compiling.

- **A variable whose name spells a contextual keyword is read as a name.**
  MS-VBAL 3.3.5.2 does not reserve `Text`, `Error`, `Object`, `Read`, `Step`,
  `Base` and the rest of that set, so `Dim Text As String` is legal VBA - but
  the reference scanner accepted only `identifier` tokens. The result was
  incoherent: `Text = 1` reported `Variable not defined`, and `Debug.Print Text`
  on the very next line reported nothing. Fourteen of the sixteen words were
  affected. Same family as the keyword-named assignment target in #46.

  Each of these words also has a position where it IS syntax, and those are now
  guarded: `On Error GoTo`, the bare `Error` statement wherever a statement
  starts, `Exit Property` and the `End Property` footer, `For ... Step`, and the
  `Open` mode, access, and lock clauses. Across the workbook corpus the change
  is finding-for-finding identical - 41 before, 41 after.

- **`Error` and `Error$` are known runtime functions.** `Error[$]([errornumber])`
  reads the message text for a code, defaulting to the current `Err.Number`. The
  table had excluded it as an MS-VBAL reserved name, which it is not.

- **A class module's name used as if it were an instance is reported** (#47).
  `Ticket.ChangeTest`, where `Ticket` is a plain class, is `Variable not
  defined` - the VBE refuses to compile it - and the analyzer said nothing,
  because the bare-qualifier check accepted any project member surface as a
  legal receiver. What actually makes a module's name usable as a value is a
  DEFAULT INSTANCE: standard modules are namespaces, documents and UserForms
  always have one, and a class has one only with `Attribute VB_PredeclaredId =
  True`. That bit is now carried from the workbook reader to the analyzer, and
  parsed from the module's own header for a standalone `.cls` export.

  It has three states and only a vouched-for **false** reports. A host that
  never read the attribute leaves it unknown, and unknown stays silent:
  measured across the workbook corpus, 44 class modules answer the question and
  12 of them (the stdVBA library in `fullBuild.xlsm`) are predeclared, so a
  guess would have turned working code red. The same bit fixes the mirror
  defect - a predeclared class read bare, `Set x = stdArray`, was being called
  undefined when it names a real value.

- **A boolean module attribute is no longer read as empty**. The `Attribute VB_*`
  reader accepted only a QUOTED value, and the VBE quotes strings while leaving
  booleans bare, so `VB_PredeclaredId = True` and `VB_Exposed = True` both
  answered the empty string. That silently disabled the document-module fallback
  in `classifyModuleType`, whose entire job is to recognise a host module by
  those two being True - the branch could never fire. Module classification is
  unchanged across the test workbooks; the fallback simply works now.

- **An unreadable form designer no longer turns its own code-behind red** (#48).
  A UserForm's controls are declared by its designer, not its text, so the host
  seeds them - and that seed has three states, not two: a list, a vouched-for
  EMPTY list, and no answer because the designer could not be read.
  `undeclared-variable` read the third as the second, so every control the form's
  own code named was reported undeclared. The VBE stops handing out a form's
  designer once the form has been shown, which means running your own form
  turned its code red against source that had compiled a minute earlier. A form
  vouched for as empty still reports, which is the case worth keeping; the member
  rule already drew this line and now both do.

- **A Function with a declared return type that never assigns it is reported**
  (#46). The rule skipped any declared return, so a `Function ... As Double`
  that never names itself - returning 0 to every caller, compiling cleanly, and
  found later by a wrong number rather than an error - said nothing, while the
  catalogue advertised exactly that check. Widening it was measured over 67
  modules of third-party code: 40 findings, almost all false, so it ships with
  the three things that made them false. A return whose FIELDS are assigned
  counts as assigned (`MakePoint.X = 1`), a class's empty body is an interface
  member it declares and does not implement, and a body whose work is to raise
  owes no return. That leaves 2 findings on the same corpus.

- **Three more rules read the statements a single-line `If` executes** (#46).
  A single-line `If` is one leaf statement, so every rule that reads a statement
  structurally saw only `If`: `argument-count`, `const-assignment` and
  `assignment-object-type-mismatch` all missed the defect when it was written
  `If ok Then ...`. Each now reads the branches after `Then` and after `Else`.
  Deliberately opt-in rather than folded into the shared walk: a rule that scans
  statement text already sees inside a single-line If, and visiting the branches
  globally made `array-subscript-out-of-bounds` report the same defect twice,
  while visiting the CONDITION as a statement read `If MAX = 10 Then` as an
  assignment to a constant. Both are pinned.

- **A return assigned under a keyword-spelled name counts** (#46). The lexer
  classifies `Text`, `Read` and `Type` as keywords, and the assignment reader
  demanded an identifier, so an assignment to a variable or Function named one
  of them was invisible: `Function text()` assigning `text = 1` reported that it
  never assigns its own return. A name that spells a keyword is still a name -
  the bare `=` after it is what settles the statement, and `Set`/`Let` are
  handled before it. This sits in the shared assignment reader, so every rule
  built on it stops being blind to those names.

- **A return assigned inside a single-line `If` counts** (#46). A single-line
  `If` is one leaf statement, so the statement walk never reached the assignment
  after `Then`, and an untyped `Function` whose only assignment was
  `If ok Then F = 1` reported that it never assigns its own name. The rule now
  reads the branches after `Then` and after `Else` too. The same blind spot is
  in the shared statement walk, so other assignment-seeking rules have it; that
  is left alone here because it touches every rule at once.

- **Six agent tools said Excel when they serve every host**. `xlide_listModules`
  described itself as reading "the live Excel workbook", and `exportModules`,
  `configureExportMode`, `validateWorkbook`, `analyzeWorkbook` and
  `searchModules` all spoke of a workbook, while each takes a file path and
  works on Word, PowerPoint and Access too - `xlide_readModule` sitting beside
  them said so correctly, and the two disagreed. An agent reading the narrower
  wording would conclude it could not enumerate the modules of a `.docm` and
  reach for something else. Wording only; no behaviour changed, and the
  genuinely Excel-only tools (`listSheets`, `readCells`, `readFormulas`,
  `writeCells`) keep theirs.

## [4.1.3] - 2026-08-21

- **A repeat analysis no longer costs multiples of the first** (#45). Four
  analyzer memos find their entry by comparing a whole source string. Within one
  pass that is free: the caller hands back the very instance the entry was
  stored under and the comparison settles on the pointer. Across passes it is
  not, because a host that re-materialises module text - a worker boundary, a
  pipe, a re-read - produces a NEW string with the same content, so every lookup
  compared the whole module. `statementTokensCached` is asked hundreds of
  thousands of times per pass, which made every analysis after the first
  quadratic in module size: a 64,802-line module took 0.9s cold and 15.8s every
  time after. Each memo now adopts the caller's instance on a content hit, so
  the first lookup of a pass pays one comparison and the rest settle on the
  pointer. Measured on an 18,002-line module, a repeat analysis went from
  2,666 ms to 492 ms and the curve is linear again; findings are identical.

## [4.1.2] - 2026-08-21

- **`If cond Then: stmt` no longer owes an `End If`**. Statements were split on
  every colon, so a colon straight after `Then` left a bare `If ... Then`
  segment, which reads as a block opener. It consumed the real `End If` and the
  ENCLOSING block reported the missing one, pointing at a line that was fine.
  Tim Hall's VBA-JSON uses the idiom twice and reported two errors on code that
  compiles. A colon after `Then` now stays inside the single-line If.

- **A named argument may be named after a keyword**. `:=` has exactly one
  meaning in VBA, so the word before it is a parameter name - but the check
  demanded an identifier token, and `Type` lexes as a keyword. So
  `ThisWorkbook.BreakLink Name:="test", Type:=xlLinkTypeExcelLinks` reported a
  positional argument following a named one. The Office libraries name 385
  parameters across 370 members that way, `Type` alone on 197, with `Text`,
  `Variant`, `To` and `String` behind it.

- **An argument that takes an enumeration offers its values**. At `Type:=` the
  model knows the parameter is an `XlLinkType` and that the enumeration has
  exactly two members, but the caret got the general identifier list instead -
  every global and constant in the library, in no useful order. The accepted
  values now sort to the top, matched by NAME for a named argument and by
  position otherwise, in both the bare and `Call ...(...)` forms. 390 Excel
  parameters declare a modelled enumeration. The general list stays underneath,
  because `Type:=someVariable` is legal too.

## [4.1.1] - 2026-08-21

- **A code name paints beside `Me` again** (#44). In a document module,
  `Sheet1.Calculate` stopped being colored while `Me.Calculate` beside it kept
  its token, which is exactly the inconsistency #31 exists to prevent - a
  module's code name and `Me` are the same object. 4.1.0 rewrote the collector
  to resolve a member the way hover does, and two details of that went wrong: a
  document module's code name resolves to a combined project-and-host receiver
  whose surface is owned by the PROJECT name, so the host-member check read it
  as "not a host type"; and the collector flattened every visible project type
  to a class, so the module's own name claimed itself as one. The check now
  looks at the host half of a combined or union receiver, and the project type's
  real kind is carried through.

## [4.1.0] - 2026-08-21

- **A blank line keeps the indent the editor gave it** (#43). VS Code trims
  whitespace it inserted itself as soon as the caret leaves the line, so
  pressing Enter twice and arrowing back up landed the caret at column 1. The
  VBE does not do that, and the spaces it keeps are in the code store rather
  than a rendering artefact, so the two editors disagreed on a gesture people
  use constantly. `editor.trimAutoWhitespace` is now off for XLIDE's VBA
  modules. Backspace on a blank indented line clears the whole indent in one
  press,
  so keeping the indent does not make an unwanted blank line cost a press per
  tab stop.

- **Enumerations joined the model**. `Dim k As XlAxisType` is ordinary VBA and
  the name resolved to nothing at all: no completion, no hover, no coloring,
  across 1,744 enumerations. They now complete in a declaration type position,
  hover with the reference's description, color as enums, and answer as a
  qualifier - `XlAxisType.` offers exactly its three constants, and Option
  Explicit no longer calls the qualifier an undeclared variable. All 19,816
  enum constants carry a description: the reference's own where it has one,
  and their enumeration's where it does not.

- **A shared reference page no longer names the wrong application**. Microsoft
  publishes one page per shared object under every host namespace without
  substituting the application name, so a PowerPoint developer hovering
  `Axis.ReversePlotOrder` read "True if Microsoft Word plots data points from
  last to first". A mention is rewritten only when the whole sentence appears
  verbatim in the named application's own library, which is what proves it was
  cross-published; a sentence genuinely about another application - a chart's
  data really does live in "an external Microsoft Excel workbook" - is left
  exactly as written. 16 descriptions repaired, every cross-application
  reference kept.

- **Dropping into a `With` block opens the completion list**. Typing a dot
  opens it because `.` is a trigger character, but the dot Smart Enter inserts
  is not typed, so the caret landed after a dot with no list and backspacing
  over it and retyping the same character was the only way to get one.

- **The object models describe themselves, everywhere**. The type
  libraries state a type for every property and Microsoft's reference
  describes most members, but almost none of that reached a tooltip: a
  property hovered as a bare `Range.Value` with no type and no read/write
  contract, an enum member showed its value and nothing else, and a type
  name in a `Dim` hovered with no prose at all. Across Excel, Word,
  PowerPoint and Access, all 34,241 members now carry a declared type or a
  call signature, 21,249 carry the reference's own description and the rest
  a note composed from the declaration and marked as derived, 15,570 of
  19,816 enum constants carry their documented meaning, and 1,470 types
  carry theirs. Parameter descriptions reach signature help in the
  generated hosts, which used to drop them.

- **A member chain no longer dies where the type library says `Object`**
  An accessor that hands back a real object is declared `As Object`
  in COM, so `.Chart.Axes(xlCategory).HasTitle` resolved in Excel - whose
  model had the return types transcribed by hand - and nowhere else. The
  repair is now derived from the reference corpus by rules shared across
  every generator, validated against all 85 return types Excel had
  transcribed, plus a small table for the facts no page states. 6,881
  members now name a type a chain can follow. The shared Office library
  binds to its host, so `Shape.Application` reaches Excel's Application in
  Excel and Word's in Word.

- **Host members are colored, at any depth and inside `With`**.
  Coloring reached a method on a chain root and stopped: `.HasChart` inside
  a `With` block, and every property anywhere, stayed plain. A member is now
  colored exactly when hover can describe it, methods and properties told
  apart, with the same refusal to guess - an unresolved hop ends the chain
  and everything after it stays plain.

- **Removed 28 members Excel does not have**. Hand-written entries had
  been copied onto the wrong type: `TextFrame.HasText` belongs to
  TextFrame2, `WebOptions.AlwaysSaveInDefaultEncoding` to DefaultWebOptions,
  `PivotTable.PivotFilters` to PivotField, `Application.IconSets` to
  Workbook. `TimelineState.SetFilterDateRange2` existed nowhere at all.
  Completion offered every one of them. Each removal was checked against the
  installed Excel type library and the published reference, and a test now
  holds the model to what the type library actually carries.

- **Painting a module of `With` blocks got 95x faster**. Resolving a
  leading-dot member re-lexed the whole enclosing procedure, so cost grew
  with the square of the procedure's length: 546 ms to paint 400 `With`
  blocks, against 3 ms for the same code written any other way. The scan now
  reuses the module's token stream and is indexed once per procedure - the
  same file paints in 5.7 ms, and a 15,000-line module in 63 ms.

## [4.0.7] - 2026-08-20

- **Installing the test support module no longer re-cases your project**
  (#38). VBA cases identifiers project-wide to the latest declaration it
  sees, and XlideAssert declared `ByVal value`, `message`, `condition` and
  `macroName` - so installing it re-spelled every `Err.Number`, `.Value`
  and `Message` in the workbook to lower case, permanently. Every name the
  injected modules declare now carries the canonical casing of the host or
  runtime name it shadows, measured against the object models rather than
  a hand-kept list, and a test enforces the rule mechanically. The runner
  and dispatcher modules got the same treatment (their JSON wire keys are
  unchanged).

- **A module someone else re-cased no longer reads as outdated** (#38).
  The version gate compared case-sensitively, so a project whose own
  declarations re-cased XlideAssert reported a freshly installed module as
  needing an update. Identifier case now folds in that comparison - VBA
  owns it - while string literals and comments still compare verbatim, so
  a real edit still shows. The module carries a revision stamp, so a
  future fix that changes only casing still reaches installed copies.

- **A Private test Sub is reported, not compiled into a broken run**
  (#39). Discovery accepted any zero-argument Sub, but the generated
  runner calls tests as `Call Module.Proc`, which cannot compile against a
  Private or Friend target - so one `' @xlide-test` over a Private Sub
  failed the whole generated module and with it every test in the run.
  Discovery now skips them, matching the dispatcher's existing filter, and
  the directive gets a diagnostic saying the Sub must be Public.

## [4.0.6] - 2026-08-19

- **A real VBE export reads unchanged too** (#37). The VBE's Export writes
  the module text plus one trailing CRLF the live module never carries, so
  after #36 a form off a real export still read "Will update" by one
  phantom blank line. The preview strip now trims the trailing blank-line
  run the way it always trimmed the leading one - status, diff, and the
  form compare all see the same text - and interior blank lines stay
  untouched.

## [4.0.5] - 2026-08-19

- **A form can read unchanged in the sync plan** (#36). The compare held
  the `.frm` file - designer header and all - against live module text,
  which reads equal only when the engine can compose the full form export;
  any form whose designer cannot be read (and any `.frm` the VBE itself
  exported) compared unequal forever, a "Will update" that never cleared.
  Forms now compare on the half their module text can say - the same
  header-stripped text the plan's default diff already shows, so the
  status and the diff always agree - in both the import and export plans.
  Standard and class modules keep the raw comparison: their headers
  round-trip, so an attribute edit stays a real pending change. A
  designer-only repo edit reads unchanged by this rule (the designer still
  travels whenever the row is applied); a deeper designer compare remains
  open by choice.

## [4.0.4] - 2026-08-19

- **Host constants paint in one tier** (#35). The grammar could only carry
  a curated nineteen Excel names, so `xlUp` wore constant blue while
  `xlLandscape` sat plain on the same line - in every host. Resolved host
  constants now take a semantic `enumMember` token from the module's own
  model (`wdAlignParagraphCenter` and `msoTrue` in Word, `ac*` in Access),
  under the same shadow and position gates as the injected globals; the
  grammar's static list stays as the offline fallback.

- **A declared local with a host type paints its method calls** (#33).
  `Dim rng As Range` then `rng.InsertParagraphAfter` hovered "Word host
  method" while painting plain - declared names were only ever shadows to
  the paint. The collector now types declared locals and parameters from
  their `As` clauses: a name typed the same everywhere in the module
  paints its resolved host methods, a name that is untyped, ambiguous
  across procedures, or shadowed by a project type stays plain.

- **The host Global interface resolves everywhere** (#34). Word's
  `InchesToPoints`, Excel's `Union` and their siblings - the members VBA
  calls bare - answered nothing on any surface. Each model now names its
  hidden Global interface (Excel's newly promoted from the reference
  corpus) and every surface consults it: hover and completion answer with
  the verified signature and documentation, the call tip works inside a
  bare call, the undeclared and unknown-call rules know the names, member
  access chains through typed returns (`Union(a, b).Select`), and the
  paint shows a bare Global method as a call and a Global property as a
  host value. The VBA runtime's own names keep winning bare resolution
  first, exactly as VBA binds them.

## [4.0.3] - 2026-08-19

- **Member access follows a control member into the object it returns**
  (#32). `Views.SelectedItem` always hovered `As Tab`, but completion and
  hover refused the second hop into the returned object - the one receiver
  shape that dead-ended while host chains resolved. The chain now steps
  through a control member's returned MSForms type exactly when the forms
  metadata carries that surface, and the metadata now includes the
  returned-object types the library defines as interfaces - Tab, Tabs,
  Pages, Controls, Font - pulled in by reachability from the control
  members that return them. `Views.SelectedItem.Caption` completes and
  hovers as `Tab.Caption As String`, `ViewNote.Font.Bold` chains the same
  way, and a primitive return still ends the chain instead of guessing.

## [4.0.2] - 2026-08-19

- **The paint's shadow rule reaches the receiver too** (#30). A form control
  named exactly like a host global no longer wears the global's tint inside
  its own form: the global collector now takes the designer-declared
  controls the member collector already honors, so both halves of the name
  agree that the control wins the binding.

- **`Me.` paints like its code name** (#31). In a document module,
  `Me.Calculate` and `Sheet1.Calculate` are the same call, and hover,
  completion, and signature help already said so - only the paint
  disagreed. The host method collector now takes the module's `Me` host
  type (Excel.Worksheet/Workbook/Chart, Word.Document) and paints
  Me-qualified resolved methods under the same conservative gates;
  form `Me` stays with the MSForms collector, and longer chains and
  With-block members stay out of scope.

- **The shared Office constants reach every host.** Every Office VBA project
  auto-references the Office object library, but only Excel's model merged
  its enum constants - so in a Word, PowerPoint, or Access module `msoTrue`
  was invisible to hover and completion and, under Option Explicit, flagged
  as an undeclared variable. All three generated models now lay their own
  library's constants over the shared Office table (same-name chart enums
  keep their identical values), and a new round-trip suite walks every
  repo-local reference dump to prove each model resolves every constant
  with the dumped value and enum type and carries nothing invented. The
  corpus is deliberately uncommitted, so those checks run exactly where
  regeneration can happen and skip in CI and fresh clones.

- **An agent review survives the agent's second change.** The Keep/Revert
  review froze its after-image at the write that opened it, so when a second
  agent change arrived through another surface - Copilot editing the open
  module document, an editor save, a sidebar write - Revert refused with the
  drift warning and nothing captured the new state. Every XLIDE write path
  now keeps a pending review tracking the live content: Revert stays offered
  and still restores the pre-agent original the diff shows it discarding, a
  write that lands back on that original resolves the review outright, and
  the drift refusal remains only for content changed outside XLIDE entirely.

## [4.0.1] - 2026-08-19

- **Origin labels name the module's host, not Excel** (#28). Since 4.0.0
  the resolvers answer from the module's own host model, but every origin
  label stayed a literal: Word's `ActiveDocument.FitToPages` hovered as
  "Excel host method" over Word-correct data. The host model now carries
  its application name and the labels build from it - "Word host global",
  "Word host type", "Word type" in type completion, "Word/Office constant"
  with matching documentation, and a code name's completion detail names
  its real document type ("Document object" for ThisDocument, and
  "Workbook object" for ThisWorkbook instead of the old blanket
  "Worksheet object"). Excel wording is byte-for-byte unchanged.

- **A resolved host method call paints as a call** (#29). Issue #20's
  convention - a resolved method paints `function`, a property or an
  unresolved member stays untouched - covered a form's designer-declared
  controls but never host receivers, so `RegionPick.AddItem` painted while
  `ActiveSheet.Calculate`, `Application.Quit`, and Word's
  `ActiveDocument.FitToPages` sat in property blue. A host-receiver
  collector now paints them under the same conservative gates: the
  receiver is a host global or document code name at the chain root,
  shadowed by nothing in the module (designer-declared controls included),
  and the member resolves to a method on the receiver's host type. Longer
  chains and With-block members stay out of scope, exactly as they do for
  controls.

- **.xlsb answers like the other hosts instead of erroring raw.** The
  binary workbook passed the Excel gate but keeps its workbook part as
  binary `xl/workbook.bin`, so the sheet and cell tools surfaced a raw
  "Entry not found: xl/workbook.xml" and file info threw outright instead
  of answering the modules. Sheet reads and writes now refuse with the
  honest reason (save as .xlsm to use the sheet tools; VBA editing is
  unaffected), and file info answers modules and protection facts the way
  Word, PowerPoint, and legacy .xls always did. The file-locked notice
  also stopped offering Retry on failed writes - it reverted whichever
  editor was active, which could discard unrelated unsaved edits.

- **A member a form does not have is a finding now** (#26). The VBE refuses
  `EntryForm.NoSuchControl` at compile time, and so does the analyzer: a
  form's surface - code-behind, designer controls, and the MSForms
  UserForm base - proves absence exactly when its control list is
  authoritative, meaning a designer-reading host supplied it (an empty
  list included) or the source spells the controls out in its header. A
  form whose binary designer nobody has read stays out of absence claims,
  and a real `.frm` export header that defers to its `.frx` blob proves
  nothing, by measurement. `Me` inside the form follows the same rule.
  Verified against live Excel on a real four-control form: a battery of
  twenty-five early-bound member references compiled in the VBE and
  resolved in the analyzer - including `PrintForm`, which the member list
  had knowingly omitted and which this change adds.

- **F5 runs the macro in Word and PowerPoint, not just Excel.** Run Macro
  at Cursor was Excel COM end to end, so F5 in a Word module handed the
  document to Excel with workbook-language notifications. It now has full
  parity per host: the module saves, the file reopens read-only in its own
  visible application, and the macro runs through that application's COM -
  Word through its measured plain `Module.Proc` form, PowerPoint through
  its single instance with the presentation-qualified reflection call. A
  file open for editing is refused with a clear close-and-retry message,
  Access refuses with the engine's stated p-code reason, non-Windows
  platforms open the file with run guidance, and the Excel flow is
  unchanged. Live-verified in real Word and PowerPoint: the macro executed
  in both, and a missing macro surfaced as a typed run failure.

- **New Macro-Enabled File defaults and filters fit every host.** The save
  dialog's default name was NewWorkbook.xlsm, and the dialog keeps the base
  name when the type filter changes - so a Word document was born as
  NewWorkbook.docm. The default is host-neutral (NewFile), and every file
  kind now has its own labeled filter: the dialog auto-appends only a
  filter's first extension, so the bundled Word and PowerPoint filters
  could never produce a .dotm or .potm without typing the extension by
  hand.

## [4.0.0] - 2026-08-18

XLIDE opens the workspace to VBA everywhere Office puts it. Until now the
extension read one container family; 4.0.0 reads eight and writes six, each
decided from file content rather than the extension, and each verified by
its own Office application.

### Added

- **Word, PowerPoint, and Access files open natively.** `.docm`/`.dotm`,
  `.pptm`/`.potm`/`.ppsm`, and read-only `.accdb`/`.mdb` join the Excel
  formats in the explorer, the sidebar, the editor, and the agent tools -
  and so do the template and add-in variants: `.xltm`, `.xlt`, `.xla`,
  `.dot`, `.ppam`, `.ppa` (a bare VBA-project compound file), and `.mda`.
  Legacy binary containers are read too: `.doc` and `.xls` are compound
  files whose VBA storage the parser already understood, and `.ppt` holds
  its project as a zlib-compressed storage inside the `PowerPoint Document`
  stream, located through the persist chain. Access databases are read by
  a native Jet/ACE page reader that reassembles the project's streams from
  LVAL rows and chains; no Office install is involved anywhere.

- **Writes cover every container whose write path is sound.** Word and
  PowerPoint documents save the way workbooks do: OOXML packages splice the
  vbaProject part, `.doc`/`.xls` re-serialize the compound file, and `.ppt`
  rebuilds its embedded record and shifts the persist directories, the
  user-edit chain, and the CurrentUserAtom. Every write path was certified
  live: the real application opened the written file, compiled its project,
  and exported the injected module back out. Access stays read-only for a
  stated reason - Access executes compiled p-code, so source edits there
  cannot take effect - and its explorer nodes drop the write menus while
  its modules open with the editor's read-only lock.

- **The analyzer knows which host it is in.** Word (364 types, 5,808
  members), PowerPoint (201 types, 2,658 members), and Access (188 types,
  6,526 members) object models are generated from the COM type libraries
  with Microsoft Learn documentation, and the container's host token rides
  every analysis surface: live diagnostics, file analysis, the worker
  protocol, completion, hover, and signature help. `Me` in `ThisDocument`
  is `Word.Document`. `wdMainTextStory` folds as a constant in a Word
  module and stays an unknown name in a workbook; `xlLandscape` the
  reverse. `ThisDocument` offers Word's `Document_*` event stubs, and
  Excel's `Workbook_*` stubs never appear outside Excel.

- **Unit tests run in Word and PowerPoint.** The owned read-only test host
  is parameterized per application, with each host's measured quirks
  handled: Word resolves `Module.Proc` run references and requires `[ref]`
  argument marshaling, PowerPoint cannot hide its window and opens the
  presentation windowless instead, and both keep the modal watcher,
  ownership proof, and kill-on-close job the Excel host already had.
  Because Word surfaces a run target's unhandled error as a VBE modal
  instead of propagating it, `XlideAssert.Throws`/`DoesNotThrow` now run
  targets through a staged direct-call dispatcher that reports outcomes as
  values, which behaves identically in all three hosts.

- **UserForms beyond Excel.** A form in a Word document is classified,
  its designer-declared controls feed completion, it exports as a
  `.frm`/`.frx` pair, and the designer writes back - all through the same
  native MS-OFORMS machinery, proven against a form authored by Word's own
  VBE.

- **Agent writes are reviewable.** Copilot's own keep/reject cannot see a
  change made through a tool, so XLIDE supplies the contract itself: every
  chat-driven `xlide_writeModule` write opens a native before/after diff
  beside the chat, and the XLIDE tree badges the module with inline Keep
  and Revert actions until the decision is made - no notification popups.
  Revert restores the before-image through the audited write path, removes
  a module the write created, and refuses if the module changed again in
  the meantime; a rename carries the pending review to the new name.
  Controlled by `xlide.agent.showWriteDiffs`.

- **Module names beyond the project's code page work, and no longer lose
  code.** Renaming a module to a name its project's ANSI code page cannot
  store (Cyrillic in a cp1252 workbook) used to be '?'-folded on save,
  detaching the module from its source stream - the code was gone on the
  next open. Such names are now first-class: the unicode dir records and
  the CFB stream carry the real name, the ANSI records and PROJECT stream
  hold the folded projection Office itself writes, and the explorer, the
  editor, and the agent tools all speak the real name. Cross-locale
  projects authored elsewhere (and files damaged by the old behavior) read
  correctly again. Verified against live Excel: the VBE lists the unicode
  name, runs the module, and re-saves a file the engine reads back intact.
  The one refusal left is real: two names whose folded projections collide
  would duplicate a PROJECT declaration, which Excel treats as corruption.

- **Renaming a standard module no longer touches names that only look like
  it.** The cross-module reference rewrite skipped nothing: `rs.Fields.Item`
  and a With block's `.Fields.Item` had their `Fields` rewritten whenever a
  module of that name exposed a matching member, and locals, parameters, or
  module-level variables shadowing the module name had their uses rewritten
  too. The scan now respects receiver position (dot and bang access) and
  declaration shadowing, and Find References sheds the same false positives.

- **The analyzer got measurably faster on large modules.** Profiling
  against the real-workbook corpus found the editor surfaces re-lexing text
  the caches already held: the semantic-token pass dropped from 72 ms to
  15 ms per edit on a 947 KB class module, and a full diagnostics pass on
  the same module dropped from 684 ms to 390 ms. One statement-token cache
  now serves every analyzer surface, its lookups use integer keys instead
  of allocating a string per call, and identifier-shadow checks layer each
  procedure's locals over one shared module-plus-project name set instead
  of rebuilding the union per procedure. Completion stopped re-lexing the
  cursor prefix on every keystroke too: the cursor context derives its
  tokens from the cached module stream (proven equivalent at every offset,
  line continuations included), taking the per-keystroke cost on the same
  module from 22 ms to under 2 ms.

- **Every surface speaks its file's language.** The sidebar, the module
  sync and analysis GUIs, and the command notifications said "workbook" and
  offered Excel-only actions no matter which host the selected file
  belonged to. The sidebar's File Actions now name the owning application
  (a Word document gets Open in Word through the OS association instead of
  the Excel launcher pair), diff panes title their right side "File:", and
  workbook stays only where the surface really is Excel - the launcher,
  sheets, and Excel coordination settings.

- **The test host works for names and paths beyond ASCII.** The generated
  host script is now written with a byte-order mark - Windows PowerShell
  reads a BOM-less script in the system ANSI code page, which mangled the
  staged workbook path for any Windows user name with non-ASCII characters
  before a single test could run. The staged Throws dispatcher also matches
  target names under VBA's own case rule (`Option Compare Text` with
  original-case keys) instead of mixing a JavaScript lowercase with VBA's
  `LCase$`, whose mappings disagree for names like Turkish dotted I -
  verified live: differently-cased target names dispatch in real Excel and
  unknown targets still refuse.

- **One VBA test run at a time.** The Tests panel, the command palette, and
  the agent tool share one run pipeline, and nothing serialized them:
  reopening the Tests panel mid-run reset its buttons, so a second
  overlapping run could start against a host that cannot serve two
  automation owners. A run now refuses to start while another executes,
  naming the file whose run is in flight.

- **New file creation** for Word documents/templates and PowerPoint
  presentations/templates, seeded from blanks authored by their own
  applications. Formats that cannot hold the requested content refuse with
  the reason: legacy formats, `.ppsm` slideshows, non-macro formats, and
  Access databases.

### Changed

- Commands and messages that now cover more than workbooks say "file":
  Analyze File, Open File Settings, discovery and lock messages name the
  application that actually holds the file. Excel-specific surfaces keep
  their workbook wording.
- The `.frm`/`.frx` export pair produced natively in 3.7.0 is now accepted
  by Excel itself: the VBE imports the composed pair with all controls
  live in the designer, compiles, and round-trips it back out.

## [3.8.0] - 2026-08-13

### Added

- **A method call on a form's control is colored like a method.**
  `RegionPick.AddItem` now paints the way `Len` does, instead of reading as
  a plain identifier. Only a call that actually resolves is colored: a
  property such as `Taxable.Value` is left alone, so is a member the control
  does not have, and so is a name in a module that has no such control. A
  `With` block's leading dot is a member of the block's own receiver and is
  never colored as the control on the line above.

- **A control offers the members every control has.** `SetFocus`, `Move`,
  `ZOrder`, `Left`, `Top`, `Width`, `Height`, `Visible`, `Name` and `Tag` are
  declared once on the Microsoft Forms base class rather than repeated on each
  control type, so completion, hover and the call tip had none of them:
  `NameBox.` offered `Text` but not `SetFocus`. They are all offered now.

### Fixed

- **Completion on a control no longer offers Microsoft Forms' internal
  plumbing.** Fourteen `_`-prefixed dispatch entries (`_GetGridX`,
  `_SetLeft`) were being offered on a form and its controls. VBA's own editor
  hides them, and so does XLIDE now: the member table drops from 1,014 entries
  to 904.

## [3.7.0] - 2026-08-13

### Added

- **`Me.` in a form offers what a form actually has.** It offered nothing at
  all. It now lists the controls the designer declared, the form's own
  procedures and fields, and the surface every form carries: `Caption`,
  `Controls`, `Repaint`, and the members VBA adds on top of Microsoft Forms -
  `Show`, `Hide`, `Move`, `Name`, `Tag`, `Left`, `Top`, `Width`, `Height`,
  `Visible`. Those last ones are in no type library, so they were verified by
  asking a real form in Excel which members it answers to. `Me.RegionPick.`
  chains into that control's own members, and the form's name reaches the same
  surface as `Me` does, since that is how a form's predeclared instance is
  addressed.

- **Hover and the call tip answer on a form's controls.** Hovering a control
  gives `RegionPick As MSForms.ComboBox` and says which form it belongs to.
  Hovering one of its members gives the member's real call signature -
  `ComboBox.AddItem([pvargItem As Variant], [pvargIndex As Variant])` - and
  typing that call now shows a parameter tip, where before there was no
  signature to show at all.

- **The analysis worker accepts a form's controls from its host.** XLIDE's
  analyzer also runs inside other editors, and one that reads the VBA
  designer directly knows a form's controls when nothing in the module text
  does. Both the project seed and each analysis request now carry them, and a
  changed control list re-analyzes rather than reusing the previous answer.

### Fixed

- **Correction to what 3.6.0 said about reading a form's controls.** It
  claimed XLIDE reads the control list out of the form's own header. Checked
  against a real Excel workbook, that does not hold: a UserForm exported from
  Excel stores its control tree in a binary `.frx` blob rather than in the
  text of the `.frm`, and the copy inside a workbook carries no designer
  header at all. The header format 3.6.0 reads is Visual Basic 6's. Reading
  the binary designer is separate work; until then a form's controls are known
  when a host supplies them, and code-behind in a workbook-backed form can
  still report a control as an undeclared variable.

## [3.6.0] - 2026-08-13

### Added

- **A UserForm's controls are understood as members of the form.** A form's
  controls are declared by the designer, not by any line of code, so
  code-behind saying `RegionPick.AddItem "North"` is correct VBA - but every
  such reference was reported as an undeclared variable, five findings on a
  small real form. XLIDE now reads the control list out of the form's own
  header, so those references resolve, and `RegionPick.` offers that control's
  members: a ComboBox offers `AddItem` and `ListIndex`, a TextBox offers
  `Text`. A name in a form that is not a control is still reported.

### Fixed

- **No more false "object variable is Nothing" on a `Dim x As New` variable.**
  VBA creates such a variable the moment anything touches it - including after
  `Set x = Nothing` - so the error it was warning about cannot happen. The
  warning it exists for still appears on a variable that was never `Set`.

- **Completion, go to definition and signature help work on names that are not
  written in Latin script.** Four more places measured the cursor's word with
  an A-Z-only pattern: completion had no word to filter on, go to definition
  did not see a module qualifier like `Модуль.Метод`, event-handler completion
  missed the name, and signature help could not read a parameter named in
  another script.

  This closes a family of problems that also produced fixes in 3.1.5, 3.2.1,
  3.3.0 and 3.5.0. Each earlier fix stopped at the layer where the problem was
  found; the language test matrix now covers storage, analysis and the editor
  together, on Linux and Windows.

## [3.5.0] - 2026-08-09

### Added

- **Undo Rename.** Renaming a symbol edits every module that refers to it, but
  an editor's undo is per file: undoing in the file you are looking at puts
  that file back and leaves the rest renamed. A rename that spans modules now
  offers **Undo Rename**, which restores all of them together, and it is
  available from the command palette. Only the most recent rename can be put
  back, and only until something else writes to the workbook - after that the
  saved text no longer describes the file and restoring it would discard that
  change.

### Fixed

- **Expanding a large module in the explorer is much faster.** Listing a
  module's procedures worked out each one's line number by re-reading the
  module from the top, so the work grew with the square of the module size.
  Expanding a workbook containing a 26,000-line class took about 400 ms and
  now takes about 5 ms. Every workbook benefits; the large ones benefit most.

- **Pressing Enter after a `For` loop closes it when the loop variable is not
  written in Latin script.** `For Each товар In Items` closed nothing at all,
  because the check for a finished loop header only recognised A-Z names, so a
  complete header looked unfinished.

## [3.4.0] - 2026-08-09

### Fixed

- **Hiding headers now hides the whole header (issue #14).** The toggle removed
  `Attribute VB_*` lines, which is the entire header of a standard module but
  only the tail of a class or UserForm header. A class opens with a
  `VERSION 1.0 CLASS / BEGIN / ... / END` block and a UserForm with a designer
  block that nests once per control, so on a `.cls` the toggle appeared to do
  nothing and on a `.frm` it left the whole control list on screen.

- **Renaming an interface carries its member prefix (issue #9).** Renaming
  `IShape` rewrote `Implements IShape` and `Dim s As IShape` but left
  `Private Sub IShape_Draw()` behind, and a class whose prefix no longer matches
  silently stops implementing anything. Only classes that actually declare
  `Implements` for that interface are touched, and only their declarations - a
  variable elsewhere that merely starts with the same word is left alone.

- **Renaming a module name from code explains what to do.** A module's name
  lives on the component rather than in any module's text, so renaming from a
  position where only a module can stand - before a dot, or after `As`, `New`
  or `Implements` - now names the operation that does the job instead of
  reporting that the name is not renameable.

- **A rename says what it left alone.** When an unqualified call could refer to
  the symbol being renamed or to a same-named export in another module, nothing
  can prove which was meant, so it is left untouched - and now reported, with
  the module and line, because that is a decision only the developer can make.

### Added

- **Agent tools can read part of a module, search across modules, analyze one
  module, and write safely (issue #13).**

  - `xlide_readModule` accepts `startLine` / `endLine`, so an assistant working
    on a 26,000-line class no longer has to read all of it to change one line.
  - `xlide_searchModules` finds a name across every module in a workbook,
    instead of reading each one in turn. Results are capped, and the reply says
    when it stopped.
  - `xlide_analyzeWorkbook` accepts a `moduleName`, so checking the module just
    edited does not mean reading back every finding in the workbook.
  - `xlide_readModule` reports a `contentToken`, and `xlide_writeModule`
    accepts `expectedContentToken`. Because a write replaces the whole module,
    a change made by Excel or a user after the assistant's read was previously
    overwritten with no sign anything had happened. With the token, that write
    is refused instead.

## [3.3.0] - 2026-08-09

### Fixed

- **Go to definition, find references and rename now work on names that are
  not Latin.** VS Code picks the word under your cursor with a pattern that
  accepted only A-Z, so on a Cyrillic, Greek, Thai, Japanese or Chinese
  identifier it selected nothing and those commands silently did nothing at
  all. The same check rejected an `@xlide-test` procedure in a module named in
  one of those scripts before it could run.

- **Non-Latin type and interface names resolve.** `Implements` written with a
  Cyrillic interface name left the class hierarchy empty, and a variable
  declared as a class whose name is not Latin offered no members at all - not
  even the ASCII ones declared beside them. Identifier completion also dropped
  any candidate whose name was not Latin.

  This is the same family as the 3.1.5 encoding fix and the 3.2.1
  combining-mark fix, which both stopped at the analyzer. These are the
  editor-facing surfaces one layer out, and the language test matrix now
  covers them on Linux and Windows so they stay covered.

- **A rename is checked before anything is written (issue #9).** Renaming to a
  reserved word like `Sub` or `End`, or to a name already declared where the
  old one is, produced a project that no longer compiles and reported success.
  Both are refused now, with a message naming the clash. A local may still
  take a name the module declares, because VBA allows that.

- **A malformed procedure header no longer reports a second, invented problem
  (issue #10).** `Public Sub 1Bad()` reported both the bad name and an
  "'End Sub' has no matching 'Sub'" on a line where nothing is wrong. Only the
  real problem is reported; a genuinely orphaned `End Sub` still is.

- **A user-defined Type is no longer shadowed by a same-named class from a
  referenced library (issue #11).** Declaring your own `Point` offered 37
  members of Excel's chart Point instead of your fields. Your project's
  declarations are resolved first, as VBA does. Enum names also reach their
  constants now, so `Corner.TopLeft` completes.

### Added

- **Unqualified calls VBA refuses to compile are now reported (issue #12).**
  When two modules export the same public procedure name, a call to it from a
  module declaring neither is an "Ambiguous name detected" compile error - and
  read as clean. It is reported at the call site, since exporting a name twice
  and qualifying the calls is ordinary VBA. Verified against the VBE in both
  directions.

## [3.2.1] - 2026-08-03

### Fixed

- **Identifiers written with combining marks are no longer split in half
  (issue #8).** Thai writes a letter as a base plus a tone mark, and
  Devanagari as a base plus a matra. Those marks are not letters, and every
  identifier pattern stopped at the first one - so a variable you had
  declared was read as two names and reported as undefined ("Variable not
  defined: 'า'"), and a procedure with such a name never appeared in the
  explorer tree. The VBE compiles and runs these identifiers, so each of
  these was a false error on valid code. Marks now continue a name wherever
  identifiers are recognised, while still being rejected at the start of one.

  This is the same family as the 3.1.5 encoding fix, one step further along:
  that release taught these patterns about non-Latin letters, which covered
  Cyrillic and Greek but not scripts that build a letter from more than one
  character. The language test matrix now carries Thai and Devanagari
  procedure names end to end, which is what would have caught this.

## [3.2.0] - 2026-08-03

### Added

- **Create Excel add-ins and true .xlsb workbooks (issue #5).** The New
  Macro-Enabled File command (formerly New Macro-Enabled Workbook) now offers
  Excel Add-In (.xlam) alongside .xlsm and .xlsb, and the format follows the
  extension you choose. An .xlam is a genuine add-in that Excel loads as one
  (ThisWorkbook.IsAddin is True), not a renamed workbook: what decides that is
  the workbook part's content type, which only an add-in template carries. The
  same applies to the agent tool xlide_createWorkbook, whose filePath extension
  now selects the format.

### Fixed

- **Choosing .xlsb in the new-file dialog produced an .xlsm.** The dialog has
  always offered .xlsb, but creation copied the .xlsm template whatever the
  extension, so the result was an .xlsm-format file under an .xlsb name, which
  Excel rejects or repairs on open. Each extension now has its own template.

### Performance

- **The Export and Import Preview no longer lag on huge modules.** Clicking a
  module in the preview rebuilt the whole diff pane; for a 26,000-line class
  that meant building 132,000 DOM elements on every click, and the rebuild ran
  even for clicks that could not change what was shown - so ticking the
  checkbox on the row already displayed cost as much as switching modules. The
  pane now keeps only the rows in view in the DOM, and skips rebuilds that
  cannot change the display. Activating that module went from about 1.6
  seconds to 4 ms, and ticking its checkbox from about 2.2 seconds to nothing
  measurable.

## [3.1.6] - 2026-08-02

### Fixed

- **Vietnamese (cp1258) text encodes correctly.** The code page stores
  accented vowels as a precomposed base plus a combining tone byte, which is
  not the character's canonical decomposition, so saving `Tiếng Việt` wrote
  `Ti?ng Vi?t`. The encoder now folds each combining mark back into the base
  until a mapped combination appears.
- **UTF-8 (cp65001) projects encode correctly.** The reverse-table encoder
  cannot express multibyte sequences, so non-ASCII text in a UTF-8 project
  was replaced with `?` on save. UTF-8 projects now encode directly.

### Added

- **A CI language matrix guards every supported code page.** One
  native-language sample per page - Thai, Japanese, Simplified and
  Traditional Chinese, Korean, Cyrillic, Greek, Turkish, Hebrew, Arabic,
  Vietnamese, Baltic, KOI8-R/U, Mac Roman, ISO-8859, and UTF-8 - each run
  through the real engine end to end (write, read, list, validate) on both
  Linux and Windows, with native-language module names exercised for
  Cyrillic, Japanese, and Chinese, and a full-ICU guard so degraded text
  decoding can never silently pass CI. Writing this matrix is what caught
  the two fixes above.

## [3.1.5] - 2026-08-02

### Fixed

- **Non-English text is no longer corrupted (issue #6).** The 3.x native
  engine decoded and encoded every VBA project as Windows-1252, so a Russian
  (cp1251) project read 'Модуль' back as 'Ìîäóëü' - and text saved while
  displayed that way was written to the workbook corrupted. Text now converts
  in the project's actual code page in both directions, covering Cyrillic,
  Central European, Greek, Turkish, Hebrew, Arabic, Baltic, Vietnamese, Thai,
  KOI8, and the double-byte Japanese, Chinese, and Korean pages - each pinned
  by round-trip tests, plus an end-to-end cp1251 workbook fixture verified
  down to the stored bytes and opened in real Excel.

  If you saved a module through 3.0.0-3.1.3 while its text displayed
  corrupted, the corruption was written into the workbook and cannot be
  reconstructed - restore those modules from a pre-3.x copy. Modules that
  were only opened and read were never modified.

- **Procedures with non-ASCII names appear in the explorer tree.** `Sub
  Проверка()` was invisible to the procedure list, which matched identifiers
  with an ASCII-only pattern.

- **No more false errors on non-ASCII procedure names.** The structural
  analyzer could not see a Cyrillic-named Sub as a block opener, so its `End
  Sub` reported "no matching 'Sub'" on perfectly valid code. Clean Cyrillic
  and Greek modules now analyze with zero diagnostics.

### Also in this release

- A class or UserForm named directly as a receiver (factory-style
  `VB_PredeclaredId` classes) offers its members in completion, and hover
  works on any object-module surface used as a bare receiver.
- The evidence corpus is plain ASCII again, with a test keeping smart
  punctuation out of the audit and oracle files that downstream ports vendor
  verbatim and checksum.

## [3.1.3] - 2026-08-02

### Fixed

- **Re-running Analyze Workbook on an unchanged workbook is now instant, and
  runs no longer intermittently stall part-way.** Each completed analysis
  caused the explorer to fire tree updates a couple of seconds later
  (protection probes), the results panel treated those as a reason to
  re-analyze in the background, and the next explicit run then queued on the
  analysis worker behind that redundant work - stalling at the largest module,
  sometimes for several seconds. Two changes: analysis results are now cached
  per workbook against a content fingerprint (open-editor edits included) plus
  the severity settings, so an unchanged workbook answers in milliseconds with
  "Analysis up to date" instead of re-analyzing; and the results panel no
  longer re-analyzes on tree events at all, since protection badges do not
  change analysis inputs. Editing any module or changing severity settings
  still re-analyzes, incrementally.

## [3.1.2] - 2026-08-01

### Fixed

- **Analyze Workbook's progress notification counts up instead of appearing
  stuck.** Since 3.1.1 moved per-module analysis onto the worker thread, every
  per-module progress report fired within the first few milliseconds and the
  100ms display throttle dropped them all - so the toast sat on "Reading VBA
  modules..." for the whole run, which read as a hang even though the analysis
  itself is as fast as before (and roughly twice as fast on unchanged re-runs).
  Progress now reports each module's completion - "Analyzed <name> (n/N)" - so
  the counter climbs while a large module is still being analyzed.

## [3.1.1] - 2026-08-01

### Performance

- **Analyze Workbook no longer hitches the editor while it runs.** The command
  analyzed each module on the extension host - on a workbook containing a very
  large module, that meant sub-second UI stalls mid-command. Per-module
  analysis now runs on the analysis worker thread (the same path live
  diagnostics took in 3.1.0), with an identical in-host fallback if the worker
  is unavailable. Re-running the command on a workbook you have been editing is
  also faster now: the worker keeps per-module incremental state, so an
  unchanged module is not re-analyzed from scratch, and an unchanged workbook
  skips the module-source transfer entirely.

## [3.1.0] - 2026-08-01

### Performance

- **Typing in very large modules no longer lags.** The "local" diagnostics
  pass ran a full module analysis on the extension host - about 700-945 ms on
  a real 24,000-line class - 90 ms after every pause in typing, so keystrokes
  could jam behind it. Both diagnostic passes now run on the analysis worker
  thread when it is healthy: the host runs no analysis at all while you type,
  and the worker's incremental state means the follow-up pass re-analyzes only
  what changed. This also removes a similar freeze on first open of a large
  module. If the worker is unavailable, the in-host pass is paced well behind
  the typing burst instead of fired into every pause.

- **Repeated workbook reads are now nearly free.** XLIDE keeps one parsed
  workbook per file (validated against the file's timestamp and size on every
  call, so Excel, git, or another window writing the file is always seen).
  Expanding a 42-module workbook in the explorer went from paying a full parse
  on every one of its 44 internal reads to paying exactly one; warm reads
  dropped from ~4 ms to ~0.02 ms. Saves invalidate the cache atomically, and
  mutating operations never share the cached parse.

### Removed

- **Live Share integration.** Guest browsing was never functional: it required
  a shared-service API that Microsoft restricts to approved extensions, so the
  guest side always came up empty. The integration, its remote tree nodes and
  URI routing, the guest status bar item, and the `vsls` dependency are gone,
  along with the engine fallback path that existed only to serve it.

### Internal

- Retired the last bridge-era names and dead code: `BridgeError` is now
  `WorkbookEngineError`, the unused child-process watchdog option is gone, a
  full unused-export audit removed the final two dead functions, and the
  remaining Python-era test fixtures were updated. The import graph is fully
  connected and the engine's only Node dependencies are `fs`, `path`, and
  `zlib`.

## [3.0.0] - 2026-08-01

### Changed

- **XLIDE no longer needs Python. It reads and writes workbooks itself.**
  Every workbook operation used to run through a long-lived Python child
  process, which meant every install began with a setup gate: find an
  interpreter, install two libraries, and recover when either went missing.
  That whole layer is gone. XLIDE now implements the container formats
  directly - the [MS-CFB] compound file, [MS-OVBA] compression, the VBA
  project's dir/PROJECT/module streams, and the OOXML package and worksheet
  surface - in the extension itself.

  For you that means: nothing to install beyond the extension, no Setup
  section, no interpreter to configure, no backend to restart, and the
  workbook tree available the moment XLIDE loads.

  The engine was validated against the previous implementation before the
  swap: 77 modules across 4 real workbooks parse identically, compression is
  byte-identical for every module source, container round-trips are lossless
  and idempotent, and workbooks mutated through every write path open cleanly
  in real Excel with their VBA readable through the VBE.

### Performance

- **Reads are about twice as fast, and saves nearly twice as fast.** Opening a
  VBA project used to decompress every module even when the caller only wanted
  names, types, or one module out of forty - about three quarters of the cost
  of every read. Module source is now decompressed on first use, and module
  classification reads only the header. Saving re-deflates the rebuilt VBA
  project at a compression level tuned for the write path, which costs about
  half the time for a workbook well under a percent larger.

  On a 2.4 MB workbook with 42 modules: listing modules 8.1 ms -> 4.0 ms,
  reading a module 8.5 ms -> 4.0 ms, saving an edited module 34.5 ms -> 17.0 ms,
  creating one 61.5 ms -> 37.2 ms.

### Added

- **New analysis rule: Friend members reached through a late-bound receiver.**
  Friend members are not on a class's dispatch interface, so reaching one
  through a `Variant` or `Object` receiver - including a `Collection` element,
  since `Collection.Item` returns Variant - raises run-time error 438. VBA
  compiles it without complaint, so the failure only appears on the first
  execution that reaches the call, which is how this survives in shipped code.

  The rule reports only names that resolve exclusively to Friend members of
  project classes, and stays silent when the same name is Public anywhere, is
  part of the Excel object model, or belongs to a VBA runtime object - the
  runtime type inside a Variant is unknowable, so anything less certain would
  be a guess. Backed by three Excel/VBE oracle cases, including the control
  proving that the same read through a typed local runs clean.

  Scanning five real workbooks with it turned up three genuine instances of
  the bug in one function.

### Fixed

- **The VBA test host now proves it owns its Excel instance.** It quits that
  instance when a run ends and kills it on a hang, so attaching to an Excel you
  already had open would put your unsaved work at risk. It now snapshots the
  running Excel processes before creating its own and refuses the run rather
  than touching one of yours. The owned instance is also tied to the host
  process, so a crashed run cannot leave a hidden Excel holding your workbook
  open.

- **Fewer ways for a test run to wedge on a dialog.** Macro security is
  suppressed for the host's own instance (that prompt is Excel-owned and cannot
  be dismissed by anything), alert suppression is re-asserted before closing and
  quitting in case test code turned it back on, and the modal watcher now covers
  opening and teardown instead of only macro execution.

### Removed

- The `xlide.pythonPath` and `xlide.checkPythonLibraryUpdates` settings, the
  Setup section of the sidebar, the Python setup commands, and the
  `xlide_runOpenpyxl` agent tool. Sheet and cell operations are covered by
  `xlide_readCells`, `xlide_readFormulas`, and `xlide_writeCells`.

### Notes for AI agent users

- Tool descriptions now state plainly that VBA lives in the workbook, reachable
  either through the XLIDE tools or the `xlide-vba://` virtual file system where
  saving a module writes straight into the workbook - and that exported
  `.bas`/`.cls`/`.frm` files in a project are one-way artifacts that no workbook
  reads back, so editing one changes nothing. Agents should leave them alone
  unless you explicitly ask about the export artifacts.

## [2.6.1] - 2026-07-25

### Fixed

- **Workbook protection/signature detection works with pyOpenVBA 3.1.0.**
  pyOpenVBA 3.1.0 removed an import alias XLIDE relied on, so the signed/locked
  workbook badges and the `getProtectionInfo` agent tool failed for fresh
  installs (and for anyone using the new sidebar Update button). XLIDE now
  imports from the API's canonical home, verified against both pyOpenVBA 3.0.1
  and 3.1.0.

## [2.6.0] - 2026-07-25

### Performance

- **Very large modules are now fast - roughly 100x less editor-thread work.**
  A real-world 24,000-line class module used to cost ~10 seconds of analysis on
  the editor thread per pass, re-running after every change; typing, completion,
  and hover all contended with it. Three layers of work landed:
  - The analyzer itself is ~9x faster: several quadratic hot paths (conditional-
    compilation activity, doc-comment extraction, per-statement re-lexing,
    symbol and type-name lookups) were replaced with one-pass indexes and
    shared-token derivation - verified byte-identical diagnostics across three
    real-world codebases at every step.
  - Editing only a procedure body now re-analyzes just that procedure and
    splices the cached results for the rest of the module, with automatic
    full-pass fallback on any declaration/signature/directive change.
  - Full analysis passes run on a background worker thread, so even the first
    full pass of a huge module never blocks typing. If the worker cannot start,
    analysis transparently falls back in-process.

### Added

- **The sidebar offers a one-click Update when the required Python libraries
  are outdated.** Once per session (opt-out: `xlide.checkPythonLibraryUpdates`)
  XLIDE compares the installed pyOpenVBA/openpyxl versions against PyPI; when
  one is behind, the Required Python Libraries row shows a yellow Update
  Available state whose button upgrades both and restarts the backend.
  Outdated-but-installed libraries never block any feature.

### Fixed

- **Opening a large class module no longer flashes hundreds of bogus "'Me' is
  only valid in a class..." errors.** Passes that ran before the workbook
  reported the module's kind analyzed it as a standard module; they now wait
  and retry instead of publishing wrong-kind errors, and the fast per-keystroke
  pass reuses the last known module kind.

### Changed

- **Fewer popup notifications.** Success summaries whose outcome is already
  visible elsewhere (analysis results, validation passed, export/import
  success, "copied to clipboard") now show as a transient status-bar message
  instead of a bottom-right toast. Failures and actionable prompts are
  unchanged.

## [2.5.13] - 2026-07-09

### Fixed

- **The workbook tree recovers after a transient read failure instead of
  staying broken until a window reload.** If a tree expansion failed - most
  commonly because Excel briefly holds an exclusive lock on the workbook while
  opening or saving it - the workbook node was left permanently empty, and even
  closing Excel did not bring it back. A failed module or procedure listing now
  shows an in-tree "Load failed - click to retry" item whose tooltip carries the
  actual error (plus a hint about Excel locks); clicking it retries just that
  listing without collapsing the rest of the tree.

_Note: v2.5.12 was tagged but never published; its fixes ship with this release._

## [2.5.12] - 2026-07-08

### Fixed

- **Two analyzer false positives on legal VBA.**
  - A `Case Is > 5` comparison clause (MS-VBAL 5.4.2.10) is no longer reported
    as an invalid operator sequence: `Is` in a Case clause is grammar, not the
    object-identity operator. Operator runs inside a Case body still flag.
  - `s = 3000000000&"x"` is no longer reported as juxtaposed values: the VBE
    reads the glued `&` as concatenation because the digits overflow Long
    (oracle-verified accepted), so a `&`-suffixed integer literal never
    provably ends a value. The in-range form (`n = 5& 1`) is deliberately
    under-reported rather than risking the false positive; other suffixes
    (`n = 5% 1`) still flag.

### Added

- **A corpus-wide no-false-positive sweep over the accepted oracle cases.**
  Every VBE-verified accepted case in the corpus is analyzed with its full
  cross-module project context; any compile-error diagnostic (or runtime-error
  diagnostic on a runtime-verified case) firing on one now fails the test
  suite. This is the gate the juxtaposition false positive slipped past.

## [2.5.11] - 2026-07-08

### Fixed

- **Five analyzer false positives found by analyzing a real-world JSON library.**
  Valid VBA that the VBE compiles is no longer flagged as an error:
  - `Open path For Binary Access Read As #f` no longer reports the `Access`
    keyword as an undefined variable.
  - Assigning a Byte array to a String (`s = bytes`, or a `Function ... As
    String` returning its Byte-array buffer) - VBA's documented encoding
    conversion - is no longer reported as an array-to-scalar error.
  - `If n > 0 Then ReDim a(1 To n)` no longer reports the ReDim's own target as
    an unallocated-array access (single-line `If` Then/Else arms are now
    recognized as allocation sites).
  - `ReDim rows(1 To n) As Collection` no longer reports the type name in the
    `As` clause as an undefined variable.
  - A parenless call whose first argument is parenthesized - `AssertTrue
    (cond), "message"` - now counts all its arguments instead of reporting a
    wrong argument count.
  In every family the neighboring genuine errors still flag (e.g. `x = Access`,
  a Long array assigned to a String, real unallocated accesses, undeclared
  ReDim bounds, and truly missing arguments).

- **macOS Python setup now offers Download and detects an install
  automatically.** On a Mac without Python, macOS's built-in `python3` stub made
  the setup failure unrecognizable, so the sidebar offered only "Set Path" and
  never noticed a subsequent install. The stub is now classified correctly: the
  sidebar shows a Download button (python.org) with a plain-English explanation,
  re-checks every few seconds so a new install is picked up without a window
  reload, and well-known install locations (Homebrew, python.org) are probed
  directly - Homebrew's directory is often not on VS Code's PATH. An explicit
  `xlide.pythonPath` and a workspace `.venv` still take precedence; Windows and
  Linux behavior is unchanged.

## [2.5.10] - 2026-06-30

### Fixed

- **Indexed collection access no longer reports a false type mismatch.** Assigning
  a single element pulled from a collection - `Set ws = ThisWorkbook.Sheets("x")`,
  `Set c = ws.ChartObjects(1)`, `Set pf = pt.PivotFields(1)`, `Set s = ch.SeriesCollection(1)`,
  and the rest of the `Collection(Index)` family - resolved to the *collection*
  type instead of the element and was wrongly flagged `assignment-object-type-mismatch`.
  These now resolve to the element type (and member completion / chaining works
  through them, e.g. `ws.ChartObjects(1).Chart`), while a genuine mismatch such as
  assigning a `Range` to a `Worksheet` still fires.
- **Upgrading the extension no longer floods the Problems panel with settings
  errors.** New settings added in a release, an older workbook `.xlide_settings.json`,
  or an analysis rule code you set under a previous version that has since been
  renamed no longer blast a wall of "setting not set correctly" errors. Unset or
  defaulted values are never validated, stale workbook settings and rule codes are
  tolerated silently, and a value you genuinely set to something invalid is now a
  single gentle warning instead of a hard error on every open module.

## [2.5.9] - 2026-06-28

### Added

- **Juxtaposed expressions in an assignment are now flagged as syntax errors.**
  A statement like `n = 1 n 1` (two value expressions with nothing between them)
  no longer slips through silently - it is reported as "expected end of
  statement", matching what the VBA compiler rejects. The check is narrowly
  scoped to an assignment right-hand side, so it never fires on valid code such
  as a call written with a space (`Foo (x)`), jagged-array access (`arr(1)(2)`),
  or a type-declaration suffix (`Count&`).

### Changed

- **AI-assistant tool descriptions now keep agents on the live XLIDE module
  tree.** The VBA read/write/rename/delete language-model tools spell out that
  the workbook is a binary container, so VBA changes must go through XLIDE rather
  than editing the `.xlsm`/`.xlsb` file or its exported `.bas`/`.cls` artifacts
  directly.

### Fixed

- **Smart Tab no longer indents the whole line when you press Tab at the end of
  a line.** Tab now inserts a tab character when the cursor is at or past the
  line's content; it still indents the line when the cursor is in the leading
  whitespace, on a blank line, or when a multi-line selection is active.
- **The XLIDE Tests panel refreshes when you save a module.** Editing and saving
  a VBA module of the workbook now updates the discovered-tests view instead of
  waiting for a manual refresh.
- **A range of robustness fixes from two full-codebase adversarial review
  passes** - covering module export/sync edge cases, the Python bridge restart
  path, output-log path redaction, atomic-save file permissions, and several
  smaller correctness issues.

## [2.5.8] - 2026-06-28

### Changed

- **Live diagnostics hold "still-typing" syntax errors on the current line.** A
  syntax error on the line your cursor is on is suppressed until you leave the
  line, matching the VBE which validates a line only once you move off it - so a
  half-typed `If` no longer flashes a red squiggle while you are still writing
  it.

### Fixed

- **Member completion after a leading dot in expression position.** Inside a
  `With` block, typing a leading `.` where an expression is expected (e.g.
  `For Each wb In .`) now offers the With-target's members.
- **Resolved deferred findings from the adversarial code review.**

### Performance

- **Less redundant work on the per-keystroke analyzer and diagnostics hot
  paths**, reducing latency while typing in large modules.

## [2.5.7] - 2026-06-27

### Added

- **A setting to turn the explorer's auto expand/collapse on or off**
  (`xlide.explorer.autoExpandCollapse`, on by default). When on, the XLIDE
  explorer reveals the active module and collapses the others as you switch
  editor tabs. Turn it off to keep the tree exactly as you arrange it - switching
  tabs and expanding nodes never auto-collapses anything.

### Fixed

- **Opening a module could intermittently fail with "command not found"** (most
  often with several workbooks open) and stay broken until the window was
  reloaded. Command and virtual-filesystem registration are now resilient to a
  partial or stale re-activation, so one failed registration no longer strands
  the rest.
- **Switching between two workbooks' tabs now reliably collapses the workbook you
  left** - a strict one-workbook-at-a-time accordion. Previously the workbook you
  switched away from could stay expanded.

## [2.5.6] - 2026-06-27

### Added

- **Excel Integration settings (`xlide.excelIntegration.*`).** A new settings
  group controls what XLIDE does when a module save, add/rename/delete, F5, or
  Open Workbook needs to write a workbook that Excel holds open: a coordination
  mode (`block`, the default and safest; `closeTracked`; `closeForce`),
  `trackOpenedWorkbooks`, `reopenAfterClose`, `reopenMode`
  (`lastState`/`readOnly`/`readWrite`), and `reopenReadOnlyAfterSave`. XLIDE can
  now gracefully close and reopen a workbook in Excel to complete a write, and
  refresh Excel's stale read-only view after a save.
- **Comment continuation on Enter.** Pressing Enter on a VBA comment line starts
  the new line with an apostrophe to continue the comment. Controlled by
  `xlide.editor.continueCommentOnNewline` (default on) and
  `xlide.editor.mirrorCommentSpacing` (default on), which lines the continuation
  up with the previous comment's text.
- **New compile-error diagnostic `if-reserved-keyword-in-condition`**
  (VBE-oracle-verified): flags a reserved keyword in a block-If condition, such
  as `If If True Then` or `If True Then Then`.
- **New compile-error diagnostic for a standalone multi-argument parenthesized
  call** used as a statement (for example `mySub("a", "b")`).
- **Python availability pulse:** after you install Python, XLIDE detects it
  without a manual reload.

### Changed

- **F2 Rename now updates call sites across modules**, not only within the
  current module.
- **Inline (ghost-text) suggestions are turned off for VBA modules** so they no
  longer take priority over XLIDE's completion list.
- **The explorer tree only auto-collapses a workbook when focus moves to a tab in
  a different workbook**, not on transient focus loss.
- **A single, clear notification on a locked save or open**, instead of stacking
  XLIDE's own warning on top of VS Code's native one.

### Fixed

- **F5 on a procedure with required parameters now explains the problem** (it
  names the required parameters and how to run it) instead of failing with an
  opaque COM error.
- **F5 rides out a busy Excel** (`RPC_E_CALL_REJECTED`), such as a `MsgBox` left
  open from a previous run, and reports a clear "Excel is busy" message if it
  persists.
- **Saving a module while the workbook is open read-only in Excel** no longer
  surfaces a raw error, and F5 on a read-only workbook no longer races its own
  background refresh.
- **Transient file-lock failures while reading now retry** (Excel briefly locks
  the file while it saves) instead of a spurious "cannot read" error.
- **Race-condition hardening** across the Excel coordination paths, the explorer
  cache, dirty-module backups, and module export/import (per-folder
  serialization).

## [2.5.5] - 2026-06-20

### Changed

- **A procedure closed with the wrong `End` keyword is now a warning, not an
  error.** VBE accepts `End Sub`/`End Function`/`End Property` interchangeably as
  procedure closers (oracle-verified: `Property Get ... End Function` compiles), so
  this is a style mismatch rather than a missing-closer compile error. The new
  `mismatched-end-keyword` warning fires on the opening line, the procedure is
  still treated as closed, and the existing quick-fix swaps the keyword to match
  (e.g. `End Function` → `End Property`). Genuinely unclosed procedures and stray
  closers remain errors (`missing-block-closer` / `unmatched-block-closer`).

## [2.5.4] - 2026-06-20

A conditional-compilation accuracy patch for the modern cross-compiler idiom that
guards twinBASIC-only intrinsics behind `#If TWINBASIC Then ...`. fastjson's
LibJSON drops from 2 reported errors to 0.

### Fixed

- **`TWINBASIC` now defaults to `False`.** It is a compiler auto-constant defined
  only by the twinBASIC compiler; in Excel VBA it is undefined (False), so its
  `#If TWINBASIC Then` branches are inactive. Those branches use twinBASIC-only
  intrinsics (e.g. `PutMemPtr`/`GetMemPtr`) that are genuinely undefined in VBA,
  so analyzing the dead branch produced spurious `unknown-call` errors. Unlike a
  genuine unprovable host flag (which stays `unknown`), `TWINBASIC`'s value in
  VBA is known.
- **A boolean `#Const` now compares equal to its VBA numeric value** (`False = 0`,
  `True = -1`). Conditions such as `#Const Windows = (Mac = 0)` and
  `#If Windows And (TWINBASIC = 0)` previously mis-evaluated (a boolean was never
  equal to `0`), which could deactivate a live branch and hide its declarations -
  surfacing as a false `unknown-call` for a procedure defined inside it. Both are
  VBE-oracle verified.

## [2.5.3] - 2026-06-20

A diagnostics accuracy patch that clears a large class of false positives on
pointer/memory-heavy VBA (e.g. fastjson's LibJSON dropped from 61 reported
errors to 2), all VBE-oracle verified.

### Fixed

- **Hidden VBA built-ins are no longer reported as undeclared.** The runtime
  catalog gained the pointer functions `VarPtr`/`StrPtr`/`ObjPtr`, the
  byte-string family `LenB`/`LeftB`/`RightB`/`MidB`/`InStrB`/`AscB`/`ChrB` (and
  their `$` variants), and the `vbLongLong` `VbVarType` constant - all real,
  always-available built-ins that VBE compiles under `Option Explicit`.
- **`scalar-redim` no longer flags a member-array ReDim.** `ReDim x.arr(...)`
  (or `x!arr(...)`) resizes the dynamic-array *member*, not the container `x`;
  the rule mistook the qualifier for the array being resized. Qualified ReDim
  targets are now skipped.

### Internal

- Added four VBE-oracle fixtures (the pointer functions, the byte-string family,
  the `vbLongLong` constant, and the UDT member-array ReDim) and wired them into
  the `undeclared-variable` and `scalar-redim` audit entries.

## [2.5.2] - 2026-06-20

A diagnostics accuracy patch driven by three false-positive-shaped findings on
real stdVBA modules, each adjudicated against the Excel/VBE oracle.

### Fixed

- **`required-param-after-optional` no longer flags a Property Let/Set value
  parameter.** The mandatory value parameter of a `Property Let`/`Property Set`
  may legally follow an `Optional` index parameter (e.g.
  `Property Let X(Optional ByVal i As Long, ByVal v As Long)`); only non-value
  parameters obey the required-after-optional ordering. An interior required
  parameter after an `Optional` one is still flagged. (VBE-oracle verified.)
- **`exit-wrong-proc` no longer flags `Exit Function`/`Exit Sub` inside a
  `Property Get`.** A `Property Get` is value-returning, and VBE accepts every
  `Exit` kind there; `Property Let`/`Set`, `Sub`, and `Function` still require
  the matching keyword. (VBE-oracle verified across the full truth table.)

### Changed

- **`set-required` is reclassified from a compile error to a deterministic
  runtime error (Run-time error 91).** A bare assignment to an object target
  (`obj = ...` without `Set`) compiles cleanly and fails only when it executes -
  VBE-oracle verified across `= Null`, `= New`, and Function-return-name
  assignment. The finding itself is unchanged (it is still a real bug, e.g.
  `protGetNextDescendent = Null` should be `Set ... = Nothing`); only its evidence
  label is corrected, matching `object-variable-not-set`. The rule may now be
  downgraded to a warning via settings.

### Internal

- Added eight VBE-oracle fixtures covering the property value parameter, the
  `Exit`-in-Property truth table, and the missing-`Set` runtime-91 cases, and
  promoted the three rules to `vbe-oracle-verified` in the diagnostic influence
  audit.

## [2.5.1] - 2026-06-20

A correctness patch for the diagnostics engine.

### Fixed

- **`module-declaration-in-procedure` no longer suppresses a module's entire
  diagnostic output.** When a comment-only line was the first statement after a
  conditional-compilation directive (`#If` / `#Else` / `#End If`) inside a
  procedure body - a common 32/64-bit pattern - the rule's alternative-header
  probe tokenized the line to an empty list and threw a `TypeError`. Because
  `analyzeModule` converts any rule exception into an empty result, that single
  throw silently discarded **every** diagnostic for the affected module, not just
  the offending line. `tokenName` is now guarded against an empty token list, and
  its two near-identical copies are consolidated onto one implementation so the
  guard cannot drift again. (Behaviorally, the unguarded copy also mis-stripped an
  unterminated `[name` bracketed identifier; the unified version is correct.)

## [2.5.0] - 2026-06-17

Version 2.5.0 builds on the v2.4.0 static-analysis baseline with two goals, both
reaching their definition of done: completing the MS-VBAL §5.6 expression binder
and cashing it in for binder-dependent diagnostics, and finishing the
syntax-corpus mining so no source file is left un-dispositioned. The same
no-false-positive discipline applies - every shipped red has positive, negative,
and no-diagnostic controls plus a named evidence source (MS-VBAL, the Excel/VBE
oracle, or deterministic XLIDE metadata), and anything not provable stays quiet
and is deferred with a documented reason. The TypeScript suite grew to 2,071
tests; the Excel/VBE oracle now backs the diagnostics with 397 verified cases.
See `docs/static_analysis_completeness_2.5.0.md` for the auditable record.

### Added

- **`argument-shape-mismatch`** (compile-error): a bare array variable or
  same-module user-defined `Type` value passed where a parameter is a scalar - or
  a scalar (including `Variant`) passed where a parameter is declared an array -
  is a VBE compile error. The rule decides on declared shape only, never
  element-type coercion; it is oracle-verified across 9 cases and is disjoint from
  `byref-argument-type-mismatch`.
- **Operator-shape diagnostics**: `non-scalar-binary-operand` (an array or
  same-module `Type` used as the operand of a scalar-requiring operator),
  `is-operator-non-object` (`Is` on a provably scalar operand), and
  `typeof-is-always-false` (a `TypeOf x Is Y` that can never hold), all off the
  §5.6 expression AST.
- **Expression-AST structuring** for named arguments (`name:=expr`), omitted
  arguments, and bang (`!`) member access.

### Changed

- **Flow precision**: the shared dataflow now merges `If`/`ElseIf`/`Else` branch
  arms, so `object-variable-not-set` (Run-time error 91) and
  `unallocated-dynamic-array-access` (Run-time error 9) check accesses inside
  balanced `If` arms - conservatively falling back on any label, `GoTo`,
  `On Error`, or loop. Default-member (`VB_UserMemId`) matching is now a
  deterministic numeric parse against a named DISPID constant.
- **Syntax corpus fully mined**: every remaining `mining` source file is
  dispositioned and promoted to `reference`; zero files remain in `mining`.

### Performance

- A behavior-preserving pass on the per-edit analysis hot path: the operand rules
  share **one** expression-tree traversal per procedure body via an
  expression-visitor registry (rather than one walk each);
  `procedureHasUnstructuredFlow`, `moduleNonCallableSymbols`, and the same-module
  `Type`-name set are memoized per parse; `me-outside-object-module` rides the
  shared statement walk; and the branch-merge join drops a per-name allocation.
  The full 2,071-test suite stays byte-identical.

### Removed

- The one-time popup recommending users disable AI inline (ghost-text) completions
  for XLIDE VBA modules. XLIDE now coexists with inline suggestions - the `smartTab`
  keybinding yields to a visible suggestion so `Tab` accepts the ghost text - so
  the recommendation was obsolete.

### Deferred (documented)

- The comparison / Boolean / string-concatenation scalar-coercion matrix, Date
  coercion, and default-member-aware diagnostics - VBA coerces these at runtime,
  so no no-false-positive compile red is provable. Numeric/host boundary overflow
  and flow phase 2 (definite assignment) remain oracle- / binder-gated. See the
  completeness report.

## [2.4.0] - 2026-06-14

Version 2.4.0 is the static-analysis completeness release. It closes the
evidence-led completeness sprint: every shipped diagnostic now has positive,
negative, and no-diagnostic (no-false-positive) controls plus a named evidence
source - MS-VBAL, the Excel/VBE oracle, or deterministic XLIDE metadata - and
everything that cannot be proven without the expression binder or further oracle
mapping is explicitly deferred with a documented reason. The release ships with
an auditable completeness record (`docs/static_analysis_completeness_2.4.0.md`).
The TypeScript test suite grew to 1,954 tests; 342 Excel/VBE oracle cases now
back the diagnostics.

### Added

- **Numeric overflow diagnostics**, oracle-verified against the live Excel VBE:
  out-of-range `Long` and `Currency` literal assignments/arguments (Run-time
  error 6), and a new compile-time `suffixed-literal-overflow` for over-range
  `%` (Integer) type-suffixed literals. The `&` (Long) suffix is deliberately
  excluded because it is ambiguous with the concatenation operator.
- **Declaration and identifier diagnostics**: invalid identifier start/character,
  empty `Type`, duplicate `Type` fields, too-many-parameters, identifier-too-long,
  `Optional`/`ByVal` UDT-parameter constraints, and non-constant `Const`/`Enum`/
  parameter-default values.
- **Control-flow and statement diagnostics**: stray `Else`/`ElseIf` outside an
  `If`, duplicate `Case Else`, `Me` outside an object module, invalid assignment
  targets, `Open` missing `For`, `TypeOf` missing its operand, and impossible /
  too-many array-declaration bounds.

### Fixed

- The structural block-balance pass no longer emits phantom
  `unmatched-block-closer` / `missing-block-closer` errors when a `: Rem ...`
  comment trails a statement (`stripVba` now blanks `Rem` at any statement start).

### Changed

- Per-rule evidence audit across all 112 diagnostic codes; the syntax corpus and
  diagnostic-influence audit are now provenance-tracked and test-enforced.
- Minor live-diagnostics performance: two whole-source rules now use the shared
  cached tokenizer, removing redundant per-keystroke lexing.

## [2.3.0] - 2026-06-11

Version 2.3.0 is the audit-remediation release: roughly 205 commits of
performance, correctness, and packaging hardening across the analyzer, editor
providers, commands, webviews, Python bridge, and VSIX layout. The headline is
the VBA analysis engine, whose quadratic hot paths were rebuilt to scale
near-linearly with module size.

Internally, the release also restructured the code it hardened: the
diagnostics engine moved to a rule registry with per-family rule modules,
commands split into per-domain modules, webview HTML/CSS/JS moved to template
assets under `assets/webview/`, workbook module operations and the
per-workbook project index became shared services, and the C# and PowerShell
test-host sources were externalized to `assets/testhost/`. The TypeScript
test suite grew from 1,604 to 1,735 tests over the effort, joined by new
real-workbook Python backend tests.

### Changed

- **Analysis engine performance** - each analysis pass now lexes and parses a
  module once and shares statement tokens, signature tables, type
  environments, and member resolution across all rules, instead of re-lexing
  per statement and re-parsing the module for every dotted reference.
  Measured on a ~4,000-line module, full analysis dropped from ~115 s to
  ~0.95 s (~122x faster), and warm member completion resolves ~350x faster.
- **One shared project index per workbook** - diagnostics, completion, hover,
  signature help, navigation, and semantic tokens now read a single
  incremental per-workbook project index that folds in module changes,
  replacing four divergent caches that each rebuilt project context while
  typing.
- **Faster startup** - the Python backend now starts lazily on first use
  (expanding a workbook, an XLIDE command, or the sidebar becoming visible)
  instead of spawning in every window at activation, the Excel host object
  model builds lazily, and the extension bundle is smaller; measured window
  startup improved by about 12%. Workspaces without Excel workbooks never
  start Python.
- Rebuilt Smart Enter on the analyzer lexer's stripped-line substrate, so
  statements hidden in trailing `: Rem ...` comments no longer fool block
  auto-close, `With` continuation, or loop-iterator sync.
- Shrank the VSIX download from about 2.1 MB to about 536 KB: the marketplace
  icon went from 1.69 MB to a 28 KB 256x256 PNG, dev-only commands are hidden
  from the Command Palette outside development mode, and local development
  configuration and coverage output no longer leak into the package.

### Fixed

- Fixed silent data loss when saving a module that had also been edited in
  the Excel VBE: module file stats now derive from the real workbook file
  mtime, so VS Code's save-conflict detection prompts instead of silently
  overwriting the newer workbook copy.
- Fixed `getWorkbookInfo` and `listSheets` crashing on every workbook with
  openpyxl 3.1 and newer, where `ReadOnlyWorksheet` lost its `dimensions`
  property and `defined_names` became a dict.
- Fixed the lexer reading a file-number `#` (for example `Write #1, x`) as a
  date-literal opener, which swallowed the rest of the statement's tokens.
- Space-triggered completion no longer fires the full resolver cascade for
  spaces typed in ordinary code; `As `-type completion, `End`/`Exit`
  keywords, event stubs, and labels still work.
- Module-name input boxes now share one validator, closing the creation path
  that accepted digit-leading module names VBA rejects.
- The `xlide_createWorkbook` agent tool now refuses to overwrite an existing
  workbook instead of silently replacing it.
- A crashed Python backend now offers a Restart Backend action instead of
  asking for a whole window reload.
- Live Share guest writes are now recorded in the host's write audit and
  refresh the host's open editors instead of clobbering the guest's change on
  the next save.
- Stale dirty-module backups are now pruned on activation instead of
  accumulating in extension global storage.

## [2.1.2] - 2026-06-09

### Changed

- Refined the marketplace README positioning for new VBA programmers, students,
  experienced VBA developers, and agentic AI workflows over real workbook
  context.
- Clarified the XLIDE workflow around auto-detected workbooks, tree navigation,
  direct workbook read/write editing, local disk push/pull sync, detailed diffs,
  and committing exported modules to version control.
- Documented that import/export is workbook-scoped when multiple workbooks are
  present, including selected-workbook previews, sidecar settings, and import
  apply behavior.

## [2.1.1] - 2026-06-09

### Fixed

- Fixed module export/import preview coloring so create rows read as additions,
  delete rows read as removals, and overwrite/update rows keep modified
  treatment.
- Improved sync preview clarity with operation-specific diff titles, semantic
  status badges, a `Select Pending` action, and clearer copy-button tooltips for
  missing workbook or repo sides.
- Rewrote the README for a broader Excel-user audience, putting the main value
  proposition, getting-started path, and full cross-platform links first.

## [2.1.0] - 2026-06-09

### Added

- Added module-qualified procedure, function, constant, enum, and exported
  global resolution across diagnostics, completion, hover, signature help,
  rename, and Option Explicit checks.
- Added module-level XML documentation comment support so module summaries and
  member comments flow through IntelliSense surfaces.
- Added `XLIDE: Copy Performance Snapshot` and optional
  `xlide.performance.trace` output logging so editor, analysis, backend,
  virtual filesystem, sidebar, sync, documentation metadata, and VBA test
  latency can be diagnosed from one recent timing buffer.
- Expanded performance trace coverage across Python bridge startup/RPC calls,
  `xlide-vba://` reads/writes, workbook analysis stages, sidebar workbook
  discovery/rendering, module sync/export previews, analysis/test result
  webviews, documentation metadata reloads, and owned Excel test execution.
- Promoted a large wave of Excel host metadata for completion, hover, signature
  help, and receiver-chain inference, including WorksheetFunction, Pivot
  objects, QueryTables, chart internals, ShapeRange, comments, sort/filter
  helpers, form controls/OLEObjects, workbook connections, slicers/timelines,
  shape and chart formatting internals, conditional formatting subtypes,
  legacy drawing objects, sparklines, XML maps, publish objects, and web options.

### Changed

- Refocused and closed the Version 2.1.0 roadmap as the completed
  red-squiggle completeness sprint, with parser, binder, diagnostics, host
  metadata, rename, IntelliSense, and performance work prioritized by
  developer-experience impact.
- Moved all remaining open 2.1.0 backlog into `docs/roadmap_version_2.2.0.md`,
  with object member/event authoring continuation as the new 2.2 Priority 1.
- Aligned normal module rename behavior with the shared module-rename strategy
  so module-qualified references update consistently outside the class-module
  tree path.
- Moved user-facing documentation comment guidance into the user guides area and
  recorded internal oracle discipline for sequential, non-parallel runs.
- Improved workbook context and analysis responsiveness with bounded parallel
  module analysis, more explicit trace stages, cancellation/supersession
  handling, and cache pruning for editor project context.
- Changed XLIDE VBA editor defaults to use 4-space indentation and a quieter
  minimap/overview-ruler profile for large generated or workbook-backed modules.

### Fixed

- Hardened shared expression binding and name resolution for bare and
  module-qualified identifiers, procedure return variables, exported globals,
  constants, enum members, arrays, `With` receivers, source-shadowed runtime
  names, and host globals.
- Reduced false positives from live workbook testing around Option Explicit,
  module-qualified calls/reads, Attribute metadata placement, function return
  checks, array-return assignments, inactive `#If` branches, `VBA7`/platform
  defaults, parser-recovered procedure headers, and late-bound object/Variant
  receiver behavior.
- Expanded high-confidence red diagnostics for declaration order, parameter and
  property shapes, duplicate labels, enum ambiguity, conditional branch order,
  `For`/`Next` variable mismatches, `For Each` control/source typing,
  `ReDim`/`Erase`/`LBound`/`UBound` array misuse, deterministic runtime
  conversions, `Null`, `CVErr`, scalar/object `Set` usage, ByRef binding, and
  straight-line object-variable-not-set cases.
- Fixed source-shadowed runtime names, intrinsic/host-global resolution pressure,
  `With` receiver lookup, enum member ambiguity, and module-qualified exported
  member lookup in red-squiggle diagnostics and IntelliSense.

## [2.0.2] - 2026-06-04

### Fixed

- Fixed workbook analysis false positives in common line-numbered VBA, including
  `On Error GoTo 0`, `On Error GoTo -1`, `Erl`, and line-numbered
  `Select Case`/`Case` blocks.
- Hardened diagnostics and call analysis around numeric line labels for
  assignments, `Set` assignments, `Const` writes, procedure calls, argument
  counts, and mismatched `Exit` statements.
- Improved workbook analysis result navigation by queueing rapid row clicks and
  centering the selected finding in the opened module.

## [2.0.1] - 2026-06-04

### Changed

- Removed contribution links from the in-extension XLIDE sidebar so the product
  shell ends with Support actions.
- Moved open-source support links to the bottom of the repository README.
- Tightened VSIX packaging excludes so workbook files and test workbook folders
  stay out of Marketplace packages.

## [2.0.0] - 2026-06-04

Version 2 turns XLIDE from a workbook/module bridge into a fuller VBA
development environment for VS Code: project-aware editing, deterministic
analysis, previewable sync, workbook tests, safer workbook mutation, and
agent-verifiable workflows.

### Added

- **XLIDE Activity Bar/sidebar** - dedicated setup health, selected-workbook
  actions, global settings, support commands, and a stable workbook-action
  surface while keeping workbook/module navigation in the VS Code Explorer.
- **Workbook-wide VBA analysis** - `XLIDE: Analyze Workbook` opens a dedicated
  results UI with module grouping, severity filters, counts, copy/export
  actions, workbook/global rule settings, tracking controls, and click-through
  navigation.
- **Agent analysis verification** - `xlide_analyzeWorkbook` returns the same
  workbook analysis shape used by the UI:
  `{filePath, moduleCount, errorCount, warningCount, problems: [...]}`.
- **Deterministic diagnostic metadata** - rules now carry category,
  VBE-compile-equivalence, diagnostic kind, stable codes, and source labels so
  red diagnostics stay reserved for proven compile/runtime errors.
- **Expanded VBA diagnostics** - v2 adds and hardens high-confidence checks for
  block balance, unterminated strings, duplicate declarations, `Const`
  assignment, invalid procedure headers, unbalanced parentheses, declaration and
  call hygiene, argument counts and type mismatches, unknown calls, invalid
  declaration names, scalar member access, object/scalar `Set` usage, return
  assignment warnings, test marker syntax, and source-backed `member-not-found`.
- **Project-aware language service** - completions, hover, signature help,
  go-to-definition, find references, rename, semantic type coloring, and code
  actions now use the shared analyzer/project model where XLIDE can bind the
  target deterministically.
- **Syntax and editor hardening** - Smart Enter, block snippets, close-keyword
  completions, keyword casing, comment continuation, paired `For`/`Next`
  iterator edits, parser recovery, and MS-VBAL-backed keyword/token handling
  were expanded for common VBA editing flows.
- **Source-backed object/member understanding** - project classes, public
  fields, properties, return-name assignments, UDT fields, document/UserForm
  code names, known runtime signatures, and the first generated Excel host
  surfaces participate in completion, hover, navigation, diagnostics, and
  assignment validation where the surface is proven.
- **Documentation metadata** - inline `'''` XML doc comments and external
  `.vbref.xml` metadata enrich hover, completion, and signature help for
  source-backed and externally documented symbols.
- **Generated host/reference metadata slice** - generated Excel reference
  metadata is now used for the proven Excel `Workbook` surface, with coverage
  and provenance tracked separately from runtime extension code.
- **Workbook VBA test runner** - marked `@xlide-test` procedures run through an
  XLIDE-owned read-only Excel host, with `XlideAssert.bas` support, skip/xfail
  metadata, tag filters, current-module/current-test commands, rerun flows,
  output capture, artifacts, and `status_for_ci.json`.
- **Agent test execution** - `xlide_runVbaTests` runs the same workbook tests
  headlessly and returns `{ok, summary, artifacts, report}` for AI-agent and CI
  verification.
- **Previewable import/export sync** - bulk export/import and current-module
  export use diff previews, explicit apply steps, changed/skipped/removed/failed
  summaries, safe true-up behavior, and one workbook settings owner.
- **Workbook and global settings surfaces** - `xlide.openGlobalSettings`,
  workbook-facing analysis/sync settings, guarded severity overrides,
  untracked-rule controls, strict sidecar parsing, and provenance labels make
  settings explicit without weakening diagnostic determinism.
- **Code actions** - deterministic quick fixes/source actions cover the shipped
  safe edit set, including adding `Option Explicit`, fixing known call syntax,
  adding/removing `Set`, inserting block closers, moving misplaced `Option`
  statements, splitting local `Dim` initializers, adding suppression comments,
  creating safe private stubs, analyzing the current module, and exporting the
  current module.
- **Safety, trust, and recovery** - dirty `xlide-vba` backup restoration,
  explicit mutation/run commands, write audit summaries, workbook lock/open-state
  checks, COM timeout/cleanup handling, support bundles, copy diagnostics, and
  redacted support output make workbook operations easier to trust.
- **Performance budgets** - v2 records launch-facing latency budgets for
  keystroke diagnostics, module analysis, workbook analysis, project index
  rebuilds, and sidebar health refreshes; deeper performance hardening moved to
  the v2.1.0 roadmap.
- **User-facing documentation** - README and user guides now cover setup,
  workbook workflows, analysis/ignores, import/export sync, testing,
  automation/CI, safety/support, and the v2 feature surface.

### Changed

- Replaced the older dedicated `xlide.diagnostics.optionExplicit` model with
  guarded global/workbook analysis settings, including `visibleSeverities`,
  `untrackedRules`, and `ruleSeverityOverrides`.
- Consolidated workbook-specific configuration under
  `<workbook>.xlide_settings.json`; older workbook sidecar names are not part of
  the supported v2 settings contract.
- Moved remaining binder, designer/UserForm metadata, external metadata,
  host-metadata completeness, workbook-to-workbook transfer, and performance
  scale work to `docs/roadmap_version_2.1.0.md`.
- Kept hard object-member absence diagnostics limited to exhaustive receiver
  surfaces; incomplete host/designer/external surfaces can power completion and
  hover without inventing red `member-not-found` errors.
- Updated macro/test workflows so execution remains explicit, Windows Excel COM
  remains execution-only, and normal read/edit/analyze/sync workflows continue
  to use the Python backend.

### Fixed

- Malformed workbook settings sidecars now report explicit settings errors
  instead of being silently treated as empty defaults.
- Tightened Live Share guest messaging around the current supported behavior:
  guests can edit host-opened VBA buffers, but only the host can browse/open new
  workbook modules through XLIDE.
- Refined release docs and roadmap references so v2 is closed and forward
  scope points to the v2.1.0 roadmap and related sub-roadmaps.

## [1.0.9] - 2026-05-26

### Added

- **`xlide_listWorkbooks`** agent tool - discovers all `.xlsm`, `.xlsb`, and
  `.xlam` files in the workspace so an agent can find the target workbook.
- **`xlide_getWorkbookInfo`** agent tool - returns sheets, VBA modules, and
  named ranges in one round trip.
- **`xlide_listSheets`** agent tool - lists worksheet names and used dimensions
  for cell-range discovery.
- **`xlide_readFormulas`** agent tool - reads raw formula strings instead of
  computed values.
- **`xlide_runOpenpyxl`** agent tool - executes arbitrary openpyxl code against
  a workbook for worksheet-level automation.
- **`xlide_renameModule`** and **`xlide_deleteModule`** agent tools - expose
  existing module mutation support to AI agents.
- **`.github/copilot-instructions.md`** - canonical XLIDE agent workflow loaded
  automatically by Copilot in the repository.

### Fixed

- Clarified the `xlide_writeModule` tool description: writing a missing module
  name creates the module.
- Updated `xlide_readCells` and `xlide_writeCells` descriptions to direct
  agents through `xlide_listSheets` when the sheet name is unknown.

## [1.0.8] - 2026-05-26

### Fixed

- Refreshed the affected module's procedure list after Rename Symbol or manual
  source edits so the XLIDE Explorer no longer shows stale procedure names until
  a full refresh.

## [1.0.7] - 2026-05-26

### Changed

- Unified the XLIDE Explorer welcome message for host and Live Share guest
  sessions, including workbook discovery, Refresh, and Live Share notes.

## Earlier Development Notes

These notes predate the current v2 changelog structure and are kept for
historical context.

### Added

- Module tree accordion behavior and debounce.
- Add Class Module command and context-menu entry.
- Improved module type detection for UserForms and document modules.
- COM window restore/focus after opening a workbook or running a macro.
- Outline/breadcrumb symbol-kind polish for VBA declarations.
- Per-workbook module-list caching, filesystem watcher debounce, and RPC
  cancellation token support.
- Smoke test command, TypeScript/Python unit tests, and CI workflow.
- Protected-workbook editing, signature-invalidation notices, protection and
  signature badges, Validate VBA Project, and New Macro-Enabled Workbook.
- Early VBA snippets, enter-time block closing, status bar items, workbook-locked
  error UX, and marketplace display-name polish.

### Changed

- Reorganized context menus into create, edit, workbook, transfer, and settings
  groups.
- Split class-module creation into `xlide.newClassModule` instead of routing it
  through the generic new-module path.

---

*XLIDE follows [Keep a Changelog](https://keepachangelog.com/en/1.1.0/) conventions.*
