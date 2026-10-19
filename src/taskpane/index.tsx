import { FluentProvider, webDarkTheme, webLightTheme } from "@fluentui/react-components";
import { createRoot } from "react-dom/client";
import { App } from "./App";

/** Follow the Office theme where the host exposes it (dark Office -> dark pane). */
function isDarkOfficeTheme(): boolean {
  try {
    const bg = Office.context.officeTheme?.bodyBackgroundColor;
    if (!bg) return false;
    const hex = bg.replace("#", "");
    const [r, g, b] = [0, 2, 4].map((i) => parseInt(hex.slice(i, i + 2), 16));
    return 0.2126 * r + 0.7152 * g + 0.0722 * b < 128;
  } catch {
    return false;
  }
}

Office.onReady((info) => {
  const container = document.getElementById("container")!;
  const root = createRoot(container);
  if (info.host !== Office.HostType.Excel) {
    root.render(
      <FluentProvider theme={webLightTheme}>
        <div style={{ padding: 16 }}>Probability Step Chart runs in Excel only.</div>
      </FluentProvider>,
    );
    return;
  }
  const theme = isDarkOfficeTheme() ? webDarkTheme : webLightTheme;
  root.render(
    <FluentProvider theme={theme} style={{ minHeight: "100vh" }}>
      <App />
    </FluentProvider>,
  );
});
