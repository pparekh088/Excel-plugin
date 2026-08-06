import { createRoot } from "react-dom/client";
import { FluentProvider, webLightTheme } from "@fluentui/react-components";
import { App } from "./App";

/* global Office */

Office.onReady((info) => {
  const container = document.getElementById("root");
  if (!container) throw new Error("Missing #root container");
  const hostReady = info.host === Office.HostType.Excel;
  createRoot(container).render(
    <FluentProvider theme={webLightTheme}>
      <App hostReady={hostReady} />
    </FluentProvider>
  );
});
