import * as React from "react";
import { createRoot } from "react-dom/client";
import App from "./components/App";

const rootElement = document.getElementById("container");
const root = createRoot(rootElement);

Office.onReady(() => {
  root.render(<div>If taskpane needed, then use this one.</div>);
});

if (module.hot) {
  module.hot.accept("./components/App", () => {
    const NextApp = require("./components/App").default;
    root.render(<NextApp />);
  });
}
