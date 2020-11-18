import React from "react";
import DataGrid, {
  Column,
  Selection,
  Summary,
  GroupItem,
  SortByGroupSummaryInfo,
  FilterRow,
  FilterPanel,
  FilterBuilderPopup,
  HeaderFilter,
  Export
} from "devextreme-react/data-grid";
import service from "./data.js";
import { Workbook } from "exceljs";
import saveAs from "file-saver";
import { exportDataGrid } from "devextreme/excel_exporter";

class App extends React.Component {
  constructor(props) {
    super(props);
    this.orders = service.getOrders();
  }
  render() {
    return (
      <React.Fragment>
        <DataGrid
          id="gridContainer"
          dataSource={this.orders}
          keyExpr="AccountCode"
          showBorders={true}
          onExporting={this.onExporting}
        >
          <Export enabled={true} />
          <Selection mode="single" />
          <FilterRow visible={true} />
          <FilterPanel visible={true} />
          <FilterBuilderPopup position={filterBuilderPopupPosition} />
          <HeaderFilter visible={true} />
          <Selection mode="single" />
          <Column dataField="AccountCode" width={130} caption="Account Code" />
          <Column dataField="AccountName" width={500} />
          <Column dataField="GrossAmount" width={160} format="currency" />
          <Column dataField="JournalDate" dataType="date" groupIndex={0} />
          <Column dataField="JournalNumber" />
          <Export enabled={true} />
          <Summary>
            <GroupItem
              column="GrossAmount"
              summaryType="sum"
              valueFormat="currency"
              displayFormat={"Total: {0}"}
              showInGroupFooter={true}
            />
          </Summary>
          <SortByGroupSummaryInfo summaryItem="count" />
        </DataGrid>
      </React.Fragment>
    );     
  }
  onExporting(e) {
    const workbook = new Workbook();
    const worksheet = workbook.addWorksheet("Main sheet");
    exportDataGrid({
      component: e.component,
      worksheet: worksheet,
      customizeCell: function (options) {
        const excelCell = options;
        excelCell.font = { name: "Arial", size: 12 };
        excelCell.alignment = { horizontal: "left" };
      }
    }).then(function () {
      workbook.xlsx.writeBuffer().then(function (buffer) {
        saveAs(
          new Blob([buffer], { type: "application/octet-stream" }),
          "DataGrid.xlsx"
        );
      });
    });
    e.cancel = true;
  }
}

const filterBuilderPopupPosition = {
  of: window,
  at: "top",
  my: "top",
  offset: { y: 10 }
};

export default App;
