/**
 * Type definitions for xlsx-chart library
 * https://github.com/objectum/xlsx-chart
 */

declare module 'xlsx-chart' {
  /**
   * Chart creation options
   */
  interface XLSXChartOptions {
    /**
     * Output file path
     */
    file?: string;
    
    /**
     * Chart type
     * 'column', 'bar', 'line', 'area', 'radar', 'scatter', 'pie'
     */
    chart: string;
    
    /**
     * Titles (row labels)
     */
    titles: string[];
    
    /**
     * Fields (column labels)
     */
    fields: string[];
    
    /**
     * Data
     * { "Title1": { "Field1": 5, "Field2": 10 }, "Title2": { "Field1": 10, "Field2": 5 } }
     */
    data: Record<string, Record<string, number>>;
    
    /**
     * Chart title
     */
    chartTitle?: string;
    
    /**
     * Template file path
     */
    templatePath?: string;
  }

  /**
   * Multiple chart creation options
   */
  interface XLSXMultipleChartOptions {
    /**
     * Output file path
     */
    file?: string;
    
    /**
     * Array of chart settings
     */
    charts: {
      /**
       * Chart type
       */
      chart: string;
      
      /**
       * Titles (row labels)
       */
      titles: string[];
      
      /**
       * Fields (column labels)
       */
      fields: string[];
      
      /**
       * Data
       */
      data: Record<string, Record<string, number>>;
      
      /**
       * Chart title
       */
      chartTitle?: string;
    }[];
  }

  /**
   * XLSXChart class
   */
  class XLSXChart {
    /**
     * Constructor
     */
    constructor();
    
    /**
     * Create chart and write to file
     * @param options Chart creation options
     * @param callback Callback function
     */
    writeFile(options: XLSXChartOptions | XLSXMultipleChartOptions, callback: (err: Error | null) => void): void;
    
    /**
     * Generate chart data
     * @param options Chart creation options
     * @param callback Callback function
     */
    generate(options: XLSXChartOptions | XLSXMultipleChartOptions, callback: (err: Error | null, data: Buffer) => void): void;
  }
  
  export = XLSXChart;
}