import { renderApp } from "./ui/renderApp";
import { ensureWorkbookArtifacts } from "./workbook/artifacts";
import "./ui/styles.css";

Office.onReady(async (info) => {
  if (info.host === Office.HostType.Excel) {
    renderApp();
    try {
      await ensureWorkbookArtifacts();
    } catch (error) {
      console.error("Failed to provision workbook artifacts", error);
    }
  }
});
