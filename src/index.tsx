import { Button, Spin, TextArea } from "@douyinfe/semi-ui";
import { ITable, IView, ToastType, bitable } from "@lark-base-open/js-sdk";
import json2md from "json2md";
import React, { useCallback, useEffect, useState } from "react";
import ReactDOM from "react-dom/client";
import i18n from "./locales/i18n";

/**
 * LoadApp component for exporting table data to Markdown format.
 */
const LoadApp = () => {
  // State variables
  const [isReady, setIsReady] = useState(false);
  const [isLoadingVisible, setIsLoadingVisible] = useState(false);
  const [isLoadingSelected, setIsLoadingSelected] = useState(false);
  const [md, setMD] = useState("");

  // Initialize the component
  useEffect(() => {
    const f = async () => {
      // Fetch active table data asynchronously
      await bitable.base.getActiveTable();
      setIsReady(true);
    };
    f();
  }, []);

  /**
   * Export table data to Markdown format.
   * @param {Object} options - Export options.
   * @param {ITable} options.table - The table to export.
   * @param {IView} options.view - The view to export.
   * @param {string[]} options.recordIdList - List of record IDs to export.
   */
  const exportToMarkDown = useCallback(
    async ({
      table,
      view,
      recordIdList,
    }: {
      table?: ITable;
      view?: IView;
      recordIdList: (string | undefined)[];
    }) => {
      if (!table) {
        table = await bitable.base.getActiveTable();
      }
      if (!view) {
        view = await table.getActiveView();
      }
      // Get field metadata for the view
      const fieldMetaList = await view.getFieldMetaList();

      // Prepare table headers
      const headers = fieldMetaList.map((f) => f.name);
      const rows = [];

      // Iterate over recordIdList and fetch cell data
      for (const recordId of recordIdList) {
        if (!recordId) {
          continue;
        }
        const row: string[] = [];
        for (const fieldMeta of fieldMetaList) {
          const field = await table.getFieldById(fieldMeta.id);
          let cellString = await field.getCellString(recordId);
          row.push(cellString);
        }
        rows.push(row);
      }

      // Convert data to Markdown format
      setMD(json2md([{ table: { headers, rows } }]));
    },
    []
  );

  /**
   * Export visible table data to Markdown format.
   */
  const exportVisible = useCallback(async () => {
    setIsLoadingVisible(true);
    try {
      const table = await bitable.base.getActiveTable();
      const view = await table.getActiveView();
      const recordIdList = await view.getVisibleRecordIdList();
      if (!recordIdList) {
        return;
      }
      await exportToMarkDown({ table, view, recordIdList });
    } finally {
      setIsLoadingVisible(false);
    }
  }, []);

  /**
   * Export selected table data to Markdown format.
   */
  const exportSelected = useCallback(async () => {
    setIsLoadingSelected(true);
    try {
      const { tableId, viewId } = await bitable.base.getSelection();
      if (!tableId || !viewId) {
        await bitable.ui.showToast({
          toastType: ToastType.error,
          message: i18n.t("errorMsgGetSelectionFailed"),
        });
        return;
      }
      const recordIdList = await bitable.ui.selectRecordIdList(tableId, viewId);
      await exportToMarkDown({ recordIdList });
    } finally {
      setIsLoadingSelected(false);
    }
  }, []);

  // Render loading screen if not ready
  if (!isReady) {
    return (
      <div
        style={{
          display: "flex",
          flexDirection: "row",
          justifyContent: "center",
          alignItems: "center",
          gap: 8,
        }}
      >
        <div>
          <Spin size="middle" />
        </div>
        <div>{i18n.t("initializingText")}</div>
      </div>
    );
  }

  // Render the main component
  return (
    <div
      style={{
        height: "100%",
        display: "flex",
        flexDirection: "column",
        gap: 8,
      }}
    >
      <h4>{i18n.t("title")}</h4>
      <Button
        theme="solid"
        loading={isLoadingVisible}
        disabled={isLoadingSelected}
        onClick={exportVisible}
      >
        {i18n.t("exportVisibleButtonText")}
      </Button>
      <Button
        theme="solid"
        loading={isLoadingSelected}
        disabled={isLoadingVisible}
        onClick={exportSelected}
      >
        {i18n.t("exportSelectedButtonText")}
      </Button>
      {md && <TextArea value={md} autosize />}
    </div>
  );
};

// Render the component to the root element
ReactDOM.createRoot(document.getElementById("root") as HTMLElement).render(
  <React.StrictMode>
    <LoadApp />
  </React.StrictMode>
);
