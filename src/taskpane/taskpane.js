import { addDiagnostic, clearDiagnostics, readDiagnostics } from "../shared/diagnostics";

function renderDiagnostics() {
  const output = document.getElementById("diagnosticsOutput");
  if (!output) {
    return;
  }

  const entries = readDiagnostics();
  if (!entries.length) {
    output.textContent = "No diagnostics recorded yet. Open a new meeting, try sending it, then refresh this pane.";
    return;
  }

  output.textContent = entries
    .slice()
    .reverse()
    .map((entry) => {
      return [entry.timestamp, entry.stage, entry.details].filter(Boolean).join(" | ");
    })
    .join("\n");
}

function wireButtons() {
  const refreshButton = document.getElementById("refreshDiagnostics");
  const clearButton = document.getElementById("clearDiagnostics");

  if (refreshButton) {
    refreshButton.addEventListener("click", function () {
      addDiagnostic("taskpane-refresh-clicked", "User refreshed diagnostics.");
      renderDiagnostics();
    });
  }

  if (clearButton) {
    clearButton.addEventListener("click", function () {
      clearDiagnostics();
      addDiagnostic("taskpane-diagnostics-cleared", "User cleared diagnostics.");
      renderDiagnostics();
    });
  }
}

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    addDiagnostic("taskpane-loaded", {
      host: info.host,
      platform: info.platform
    });
  }

  wireButtons();
  renderDiagnostics();
});
