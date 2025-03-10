/**
 * xlsx-chartライブラリの型定義
 * https://github.com/objectum/xlsx-chart
 */

declare module 'xlsx-chart' {
  /**
   * チャート作成オプション
   */
  interface XLSXChartOptions {
    /**
     * 出力ファイルパス
     */
    file?: string;
    
    /**
     * チャートタイプ
     * 'column', 'bar', 'line', 'area', 'radar', 'scatter', 'pie'
     */
    chart: string;
    
    /**
     * タイトル（行ラベル）
     */
    titles: string[];
    
    /**
     * フィールド（列ラベル）
     */
    fields: string[];
    
    /**
     * データ
     * { "Title1": { "Field1": 5, "Field2": 10 }, "Title2": { "Field1": 10, "Field2": 5 } }
     */
    data: Record<string, Record<string, number>>;
    
    /**
     * チャートタイトル
     */
    chartTitle?: string;
    
    /**
     * テンプレートファイルパス
     */
    templatePath?: string;
  }

  /**
   * 複数チャート作成オプション
   */
  interface XLSXMultipleChartOptions {
    /**
     * 出力ファイルパス
     */
    file?: string;
    
    /**
     * チャート設定の配列
     */
    charts: {
      /**
       * チャートタイプ
       */
      chart: string;
      
      /**
       * タイトル（行ラベル）
       */
      titles: string[];
      
      /**
       * フィールド（列ラベル）
       */
      fields: string[];
      
      /**
       * データ
       */
      data: Record<string, Record<string, number>>;
      
      /**
       * チャートタイトル
       */
      chartTitle?: string;
    }[];
  }

  /**
   * XLSXChartクラス
   */
  class XLSXChart {
    /**
     * コンストラクタ
     */
    constructor();
    
    /**
     * チャートを作成してファイルに書き込む
     * @param options チャート作成オプション
     * @param callback コールバック関数
     */
    writeFile(options: XLSXChartOptions | XLSXMultipleChartOptions, callback: (err: Error | null) => void): void;
    
    /**
     * チャートデータを生成する
     * @param options チャート作成オプション
     * @param callback コールバック関数
     */
    generate(options: XLSXChartOptions | XLSXMultipleChartOptions, callback: (err: Error | null, data: Buffer) => void): void;
  }
  
  export = XLSXChart;
}