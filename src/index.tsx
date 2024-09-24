/* eslint-disable no-console */
import { Button, Checkbox, Spin, TextArea } from "@douyinfe/semi-ui";
import {
  base,
  bitable,
  IField,
  ITable,
  IView,
  OperationType,
  PermissionEntity,
  ToastType,
} from "@lark-base-open/js-sdk";
import ClipboardJS from "clipboard";
import json2md from "json2md";
import React, { useCallback, useEffect, useState } from "react";
import ReactDOM from "react-dom/client";
import Sizes from "./consts/Sizes";
import i18n from "./locales/i18n";

/**
 * LoadApp component for exporting table data to Markdown format.
 */
const LoadApp = () => {
  new ClipboardJS(".clipboard");

  // State variables
  const [isReady, setIsReady] = useState(false);
  const [isLoadingVisible, setIsLoadingVisible] = useState(false);
  const [isLoadingSelected, setIsLoadingSelected] = useState(false);
  const [activeTable, setActiveTable] = useState<ITable | undefined>(undefined);
  const [exportAllFields, setExportAllFields] = useState(false);
  const [totalLines, setTotalLines] = useState(0);
  const [currentLines, setCurrentLines] = useState(0);
  const [duration, setDuration] = useState(0);
  const [md, setMD] = useState("");
  const [loadingTip, setLoadingTip] = useState('00/00');
  const [isLoading, setLoading] = useState(false);
  // Initialize the component
  useEffect(() => {
    const f = async () => {
      // Fetch active table data asynchronously
      const table = await bitable.base.getActiveTable();
      setActiveTable(table);
      setIsReady(true);
    };
    f();
  }, []);

  /**
   * Check if the user has the permission to export.
   */
  const checkExportPermission = useCallback(async () => {
    const hasPermission = await base.getPermission({
      entity: PermissionEntity.Base,
      type: OperationType.Printable,
    });

    if (!hasPermission) {
      bitable.ui.showToast({
        message: i18n.t("errorMsgNoPermission"),
        toastType: ToastType.error,
      });
    }
  }, []);

  /**
   * Export table data to Markdown format.
   * @param {Object} options - Export options.
   * @param {ITable} options.table - The table to export.
   * @param {IView} options.view - The view to export.
   * @param {string[]} options.recordIdList - List of record IDs to export.
   */
  const exportToMarkDown = useCallback(
    async ({ table, view, recordIdList }: { table?: ITable; view?: IView; recordIdList: (string | undefined)[] }) => {
      try {
        setTotalLines(recordIdList.length);
        if (!table) {
          table = await bitable.base.getActiveTable();
        }
        if (!view) {
          view = await table.getActiveView();
        }
        let visibleFieldIdList: string[] | undefined = undefined;
        if (!exportAllFields) {
          visibleFieldIdList = await view.getVisibleFieldIdList();
        }
        // Get field metadata for the view
        const fieldMetaList = await view.getFieldMetaList();
        const fields: IField[] = [];
        const headers: string[] = [];
        for (const fieldMeta of fieldMetaList) {
          const field = await table.getFieldById(fieldMeta.id);
          if (!exportAllFields && !visibleFieldIdList?.includes(fieldMeta.id)) {
            continue;
          }
          fields.push(field);
          headers.push(fieldMeta.name);
        }

        // Iterate over recordIdList and fetch cell data
        const rows = [];
        let lines = 1;
        for (const recordId of recordIdList) {
          setCurrentLines(lines++);
          if (!recordId) {
            continue;
          }

          const timeStart = Date.now();
          const row: string[] = [];
          for (const field of fields) {
            const cellString = await field.getCellString(recordId);
            row.push(cellString);
          }
          rows.push(row);
          const timeEnd = Date.now();
          setDuration(timeEnd - timeStart);
        }

        // Convert data to Markdown format
        setMD(json2md([{ table: { headers, rows } }]));
      } catch (err: any) {
        await bitable.ui.showToast({
          toastType: ToastType.error,
          message: i18n.t("errorMsgExportFailed"),
        });
      } finally {
        setTotalLines(0);
        setCurrentLines(0);
        setDuration(0);
      }
    },
    [exportAllFields]
  );

  /**
   * Export visible table data to Markdown format.
   */
  const exportVisible = useCallback(async () => {
    setIsLoadingVisible(true);
    await checkExportPermission();
    try {
      if (!activeTable) {
        return;
      }
      const view = await activeTable.getActiveView();
      let recordIdData;
      const recordIdList = [];
      let token = undefined as any;
      
      setLoading(true);
      do {
        recordIdData = await view.getVisibleRecordIdListByPage(token ? { pageToken: token, pageSize: 200 } : { pageSize: 200 });
        token = recordIdData.pageToken;
        setLoadingTip(`${((token > 200 ? (token - 200) : 0) / recordIdData.total * 100).toFixed(2)}%`)
        recordIdList.push(...recordIdData.recordIds);
      } while (recordIdData.hasMore);
      setLoading(false);
      if (!recordIdList) {
        return;
      }
      await exportToMarkDown({ table: activeTable, view, recordIdList });
    } finally {
      setIsLoadingVisible(false);
    }
  }, [activeTable, exportToMarkDown, checkExportPermission]);

  /**
   * Export selected table data to Markdown format.
   */
  const exportSelected = useCallback(async () => {
    setIsLoadingSelected(true);
    await checkExportPermission();
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
      await exportToMarkDown({ table: activeTable, recordIdList });
    } finally {
      setIsLoadingSelected(false);
    }
  }, [activeTable, exportToMarkDown, checkExportPermission]);

  /**
   * Copy the Markdown text to clipboard.
   */
  const copyToClipboard = useCallback(async () => {
    bitable.ui.showToast({
      message: i18n.t("successMsgCopied"),
      toastType: ToastType.success,
    });
  }, []);

  /**
   * Export the Markdown text to a file and download.
   */
  const download = useCallback(async () => {
    try {
      const blob = new Blob([md], { type: "text/plain" });
      const downloadLink = document.createElement("a");
      downloadLink.href = URL.createObjectURL(blob);
      downloadLink.download = `table-${Date.now()}.md`;
      document.body.appendChild(downloadLink);
      downloadLink.click();
      document.body.removeChild(downloadLink);
    } catch (err: any) {
      bitable.ui.showToast({
        message: i18n.t("errorMsgDownloadFailed"),
        toastType: ToastType.error,
      });
    }
  }, [md]);

  // Render loading screen if not ready
  if (!isReady) {
    return (
      <div style={styles.loadingContainer}>
        <div>
          <Spin size="middle" />
        </div>
        <div>{i18n.t("initializingText")}</div>
      </div>
    );
  }

  // Render the main component
  return (
    <Spin tip={loadingTip} spinning={isLoading}>
      <div style={styles.container}>
        <div style={styles.titleText}>{i18n.t("title")}</div>
        <div style={styles.buttonGroupContainer}>
          <Button theme="solid" loading={isLoadingVisible} disabled={isLoadingSelected} onClick={exportVisible}>
            {i18n.t("exportVisibleButtonText")}
          </Button>
          <Button theme="solid" loading={isLoadingSelected} disabled={isLoadingVisible} onClick={exportSelected}>
            {i18n.t("exportSelectedButtonText")}
          </Button>
          <Checkbox
            checked={exportAllFields}
            onChange={() => setExportAllFields(!exportAllFields)}
            style={styles.selectorCheckbox}
          >
            {i18n.t("checkboxText")}
          </Checkbox>
        </div>

        {(isLoadingVisible || isLoadingSelected) && totalLines > 0 && (
          <div style={styles.exportingInfoContianer}>
            <div style={styles.exportingInfoText}>
              {i18n.t("exportingText")} {currentLines}/{totalLines}
            </div>

            {duration > 0 && (
              <div style={styles.exportingInfoText}>
                {i18n.t("extimatedTimeText")}
                {Math.ceil(((totalLines - currentLines) * duration) / 1000)}
                {i18n.t("extimatedTimeUnitText")}
              </div>
            )}
          </div>
        )}

        {md && (
          <TextArea
            value={md}
            autosize={{
              minRows: 3,
              maxRows: window ? window.innerHeight / 2 / 20 : 10,
            }}
          />
        )}
        {md && (
          <div style={styles.buttonGroupContainer}>
            <Button className="clipboard" theme="solid" data-clipboard-text={md} onClick={copyToClipboard}>
              {i18n.t("copyButtonText")}
            </Button>

            <Button theme="solid" onClick={download}>
              {i18n.t("downloadButtonText")}
            </Button>
          </div>
        )}
      </div>
    </Spin>
  );
};

// Render the component to the root element
ReactDOM.createRoot(document.getElementById("root") as HTMLElement).render(
  <React.StrictMode>
    <LoadApp />
  </React.StrictMode>
);

interface StyleSheet {
  [key: string]: React.CSSProperties;
}

const styles: StyleSheet = {
  loadingContainer: {
    display: "flex",
    flexDirection: "row",
    justifyContent: "center",
    alignItems: "center",
    gap: Sizes.smallX,
  },
  container: {
    height: "100%",
    display: "flex",
    flexDirection: "column",
    gap: Sizes.smallX,
  },
  titleText: {
    fontWeight: 800,
    fontSize: Sizes.medium,
  },
  buttonGroupContainer: {
    display: "flex",
    flexDirection: "row",
    gap: Sizes.smallX,
    alignItems: "center",
  },
  exportingInfoContianer: {
    display: "flex",
    flexDirection: "column",
  },
  exportingInfoText: {
    fontWeight: 800,
    fontSize: Sizes.regular,
  },
};
