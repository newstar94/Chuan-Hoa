export interface TableFormatPlan {
  tableIndex: number;
  repeatHeaderRow: boolean;
  cantSplitRows: boolean;
  alignment: 'Center' | 'Left';
  borders: {
    type: 'Single';
    widthPt: number;
    color: string;
  };
  cellPadding: {
    topPt: number;
    bottomPt: number;
    leftPt: number;
    rightPt: number;
  };
}

export class TableFormatter {
  /**
   * Tạo kế hoạch định dạng chuẩn hóa cho tất cả các bảng trong tài liệu
   */
  public static planTableFormatting(tableCount: number): TableFormatPlan[] {
    const plans: TableFormatPlan[] = [];

    for (let i = 0; i < tableCount; i++) {
      plans.push({
        tableIndex: i,
        repeatHeaderRow: true, // Lặp dòng header khi tràn trang
        cantSplitRows: true,   // Không để bảng bị gãy hàng giữa 2 trang
        alignment: 'Center',
        borders: {
          type: 'Single',
          widthPt: 0.5,
          color: '#000000'
        },
        cellPadding: {
          topPt: 3.0,
          bottomPt: 3.0,
          leftPt: 5.0,
          rightPt: 5.0
        }
      });
    }

    return plans;
  }
}
