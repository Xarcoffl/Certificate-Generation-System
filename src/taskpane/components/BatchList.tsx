import * as React from "react";
import {
  DefaultButton,
  DetailsList,
  DetailsListLayoutMode,
  DetailsRow,
  IColumn,
  IDetailsRowProps,
  IDetailsRowStyles,
  MessageBar,
  MessageBarType,
  PrimaryButton,
  SelectionMode,
  Stack,
  Text,
} from "@fluentui/react";
import { useBatch } from "../state/BatchContext";
import { Batch } from "../services/types";
import { formatDateOrdinal } from "../services/dateFormat";

export const BatchList: React.FC = () => {
  const { rootDir, needsReconnect, batches, currentBatch, chooseRootFolder, reconnectRootFolder, refreshBatches, selectBatch, loading } =
    useBatch();

  const columns: IColumn[] = [
    { key: "level", name: "Training Level", fieldName: "trainingLevel", minWidth: 140, isResizable: true },
    { key: "start", name: "Start Date", minWidth: 110, onRender: (item: Batch) => formatDateOrdinal(item.startDate) },
    { key: "end", name: "End Date", minWidth: 110, onRender: (item: Batch) => formatDateOrdinal(item.endDate) },
    {
      key: "students",
      name: "Students",
      minWidth: 60,
      onRender: (item: Batch) => item.students.length,
    },
  ];

  const renderRow = (props?: IDetailsRowProps): React.ReactElement | null => {
    if (!props) return null;
    const isCurrent = (props.item as Batch).id === currentBatch?.id;
    const styles: Partial<IDetailsRowStyles> = isCurrent ? { root: { backgroundColor: "#deecf9" } } : {};
    return <DetailsRow {...props} styles={styles} />;
  };

  return (
    <Stack tokens={{ childrenGap: 10 }}>
      {needsReconnect && (
        <MessageBar
          messageBarType={MessageBarType.warning}
          actions={<PrimaryButton text="Reconnect Folder" onClick={reconnectRootFolder} />}
        >
          Access to your previously chosen folder needs to be re-confirmed.
        </MessageBar>
      )}
      <Stack horizontal tokens={{ childrenGap: 8 }}>
        <DefaultButton text={rootDir ? "Change Root Folder" : "Choose Root Folder"} onClick={chooseRootFolder} />
        <DefaultButton text="Refresh" onClick={() => refreshBatches()} disabled={!rootDir} />
      </Stack>
      {!rootDir && <Text>Choose the folder where Batches/ will be created and stored on disk.</Text>}
      {rootDir && batches.length === 0 && !loading && <Text>No batches yet. Create one in the "New Batch" tab.</Text>}
      {rootDir && batches.length > 0 && (
        <DetailsList
          items={batches}
          columns={columns}
          layoutMode={DetailsListLayoutMode.justified}
          selectionMode={SelectionMode.single}
          onRenderRow={renderRow}
          onActiveItemChanged={(item: Batch) => selectBatch(item.id)}
        />
      )}
    </Stack>
  );
};
