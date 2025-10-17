declare namespace Excel {
  interface Table {
    /**
     * Returns the data body range or a null object if the table has no rows.
     */
    getDataBodyRangeOrNullObject(): Excel.Range;
  }

  interface Workbook {
    /**
     * Returns the active worksheet in the workbook.
     */
    getActiveWorksheet(): Excel.Worksheet;
  }

  interface Range {
    /**
     * Returns the address of the range.
     */
    getAddress(
      rowAbsolute?: boolean,
      columnAbsolute?: boolean,
      referenceStyle?: Excel.ReferenceStyle,
      external?: boolean,
      relativeTo?: Excel.Range
    ): string;

    /**
     * Returns a Sort object for the range.
     */
    getSort(): Excel.Sort;

    /**
     * Returns the sparkline groups associated with the range.
     */
    sparklineGroups: Excel.SparklineGroupCollection;
  }

  interface ConditionalFormatCollection {
    add(type: Excel.ConditionalFormatType, scope?: Excel.ConditionalFormatScopeType): Excel.ConditionalFormat;
  }
}
