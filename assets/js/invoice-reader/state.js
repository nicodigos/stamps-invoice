export const state = {
  config: null,
  graphToken: sessionStorage.getItem("graphToken") || "",
  driveId: "",
  databaseRows: [],
  databaseEtag: null,
  filteredRows: [],
  filters: {},
  pagination: {
    page: 1,
    pageSize: 10,
  },
  processed: {
    summaryRows: [],
    rawRows: [],
    pendingUploads: [],
    saved: false,
  },
};

export function syncInvoiceState(shellState) {
  state.config = shellState.config;
  state.graphToken = shellState.graphToken || sessionStorage.getItem("graphToken") || "";
  if (!state.graphToken) {
    state.driveId = "";
    state.databaseRows = [];
    state.filteredRows = [];
  }
}
