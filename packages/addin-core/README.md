# @lucamattiazzi/sommelier-addin-core

Host-agnostic controller that executes Sommelier protocol methods through an Excel adapter.



Sommelier controller defaults: 10,000 cells per read, 1,000 per mutation, and 100 retained operations.
Options `maxCellsPerRead`, `maxCellsPerWrite`, and `maxOperations` accept positive safe integers.
The current table adapter loads the whole table before slicing, so the read limit applies to its full range.

`setContext(mode, ranges?)` changes the host's context hint; pinning is not an authorization boundary.
Approvals receive bounded `before`, `after`, and `affectedCells` data for writes and content clears.
Only clear(contents) is supported. Other clear modes fail without modifying cells.
Snapshots are compared before commit and after approval. Undo compares the post-write snapshot before
restoration and refuses conflicts. Snapshots preserve values and optional formulas; custom adapters
must supply `restoreRange` to undo a formula-bearing snapshot. Formatting is not part of restoration.
Actions are rejected with HOST_BUSY while another action/approval is pending. Stopping the controller
invalidates pending approvals and discards operation history. These are optimistic safeguards, not
atomic transactions across coauthoring users.
