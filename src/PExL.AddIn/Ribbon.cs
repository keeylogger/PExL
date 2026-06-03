using System.Runtime.InteropServices;
using ExcelDna.Integration.CustomUI;
using PExL.AddIn.Interop;
using PExL.AddIn.TaskPane;
using PExL.AddIn.Templates;

namespace PExL.AddIn
{
    /// <summary>
    /// The PExL ribbon tab (quick one-click actions) plus a cell context-menu
    /// entry that opens the editor. Reserved for templates; the task pane is the
    /// authoring IDE.
    /// </summary>
    [ComVisible(true)]
    public sealed class Ribbon : ExcelRibbon
    {
        private static IRibbonUI? _ribbon;

        public override string GetCustomUI(string ribbonId) => RibbonXml;

        public void OnRibbonLoad(IRibbonUI ribbon) => _ribbon = ribbon;

        /// <summary>Supplies custom button images (the PExL brand icon) to the ribbon.</summary>
        public override object LoadImage(string imageId)
        {
            if (imageId == "pexlLogo") return BrandAssets.RibbonIcon();
            return base.LoadImage(imageId);
        }

        /// <summary>Re-evaluate the Undo button's enabled state (call after a write).</summary>
        public static void RefreshUndo()
        {
            try { _ribbon?.InvalidateControl("btnUndo"); }
            catch { /* ribbon not ready */ }
        }

        // --- editor ---
        public void OnOpenEditor(IRibbonControl control) => TaskPaneManager.ShowEditor();

        // --- decompiler: read the selected cell's formula and show it as PExL ---
        public void OnTranslate(IRibbonControl control) => TemplateRunner.RunTranslate();

        // --- documentation ---
        public void OnOpenDocs(IRibbonControl control) => TaskPaneManager.ShowDocs();

        // --- undo (our own history; Excel clears its native undo on COM writes) ---
        public void OnUndo(IRibbonControl control)
        {
            ExcelInjector.UndoLast();
            RefreshUndo();
        }

        public bool OnGetUndoEnabled(IRibbonControl control) => ExcelInjector.ActiveSheetHasUndo();

        // --- templates ---
        public void OnTemplate(IRibbonControl control) => TemplateRunner.Run(control.Id);

        private const string RibbonXml = @"
<customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui' onLoad='OnRibbonLoad' loadImage='LoadImage'>
  <ribbon>
    <tabs>
      <tab id='pexlTab' label='PExL'>
        <group id='grpEditor' label='Editor'>
          <button id='btnOpenEditor' label='Open Editor' size='large'
                  image='pexlLogo' onAction='OnOpenEditor'
                  screentip='Open the PExL editor'
                  supertip='Open the side panel where you write PExL, preview the Excel formula it compiles to, and apply it to the sheet.'/>
          <button id='btnTranslate' label='Translate Formula' size='large'
                  imageMso='ConvertTextToTable' onAction='OnTranslate'
                  screentip='Translate a formula into PExL'
                  supertip='Reads the formula in the cell you have selected and rewrites it as readable PExL in the editor &#8212; great for understanding or learning from an existing formula.'/>
          <button id='btnUndo' label='Undo PExL change' size='large'
                  imageMso='Undo' onAction='OnUndo' getEnabled='OnGetUndoEnabled'
                  screentip='Undo the last PExL change'
                  supertip='Reverts the most recent change made by a PExL tool. Excel clears its own undo history when add-ins write to the grid, so PExL keeps its own (up to 100 steps).'/>
        </group>
        <group id='grpText' label='Text'>
          <button id='tplSplitColumns' label='Split to Columns' imageMso='TextToColumns' onAction='OnTemplate'
                  screentip='Split to Columns'
                  supertip='Break the selected cell on a delimiter into separate columns. Opens the editor seeded with your cell so you can pick the delimiter.'/>
          <button id='tplCombine' label='Combine Columns' imageMso='Merge' onAction='OnTemplate'
                  screentip='Combine Columns'
                  supertip='Join several cells together with a separator (TEXTJOIN). Opens the editor seeded with your selection.'/>
          <button id='tplTrimClean' label='Trim &amp; Clean' imageMso='Replace' onAction='OnTemplate'
                  screentip='Trim &amp; Clean'
                  supertip='Writes a cleaned-up copy of the cell to your left into the selected cell: strips extra spaces and nonprintable characters (TRIM + CLEAN).'/>
          <button id='tplCsvSplit' label='CSV / Paste Split' imageMso='PasteTextOnly' onAction='OnTemplate'
                  screentip='CSV / Paste Split'
                  supertip='Spread one comma-separated cell across columns (TEXTSPLIT). Opens the editor seeded with your selection.'/>
          <button id='tplProperCase' label='Proper Case' imageMso='ChangeCase' onAction='OnTemplate'
                  screentip='Proper Case'
                  supertip='Title-cases the cell to your left (after trimming) into the selected cell (PROPER + TRIM).'/>
          <button id='tplExtractEmail' label='Email Domain' imageMso='HyperlinkInsert' onAction='OnTemplate'
                  screentip='Email Domain'
                  supertip='Pulls the part after the last @ from the cell on your left into the selected cell.'/>
        </group>
        <group id='grpShape' label='Shape'>
          <button id='tplRemoveDuplicates' label='Remove Duplicates' imageMso='RemoveDuplicates' onAction='OnTemplate'
                  screentip='Remove Duplicates'
                  supertip='Distinct values from your selected range (UNIQUE). Opens the editor seeded with your selection.'/>
          <button id='tplSortBy' label='Sort By' imageMso='SortDialog' onAction='OnTemplate'
                  screentip='Sort By'
                  supertip='Sort your selected range by a key column (SORT). Opens the editor seeded with your selection.'/>
          <button id='tplFilterRows' label='Filter Rows' imageMso='Filter' onAction='OnTemplate'
                  screentip='Filter Rows'
                  supertip='Keep only the rows of your selection that match a condition (FILTER). Opens the editor seeded with your selection.'/>
          <button id='tplTopN' label='Top N' imageMso='ChartTypeBarInsertGallery' onAction='OnTemplate'
                  screentip='Top N'
                  supertip='Take the highest N rows of your selection (SORT + TAKE). Opens the editor seeded with your selection.'/>
          <button id='tplFillBlanks' label='Fill Blanks Down' imageMso='FillDown' onAction='OnTemplate'
                  screentip='Fill Blanks Down'
                  supertip='Fill a blank cell from the value above it. Opens the editor seeded relative to your selection.'/>
        </group>
        <group id='grpLogic' label='Lookup &amp; Logic'>
          <button id='tplLookupBuilder' label='Lookup Builder' imageMso='Lookup' onAction='OnTemplate'
                  screentip='Lookup Builder'
                  supertip='Look the selected value up in another sheet with a fallback (XLOOKUP). Opens the editor seeded with your cell.'/>
          <button id='tplCheck' label='If / Check' imageMso='AdpQueryCriteria' onAction='OnTemplate'
                  screentip='If / Check'
                  supertip='Multi-branch logic without nested IFs (IFS). Opens the editor seeded with your selected cell as the subject.'/>
          <button id='tplSafeDivide' label='Safe Divide' imageMso='NumberFormatDecreaseDecimal' onAction='OnTemplate'
                  screentip='Safe Divide'
                  supertip='Divides the two cells to the left into the selected cell, showing a dash instead of #DIV/0!.'/>
        </group>
        <group id='grpAgg' label='Summarize'>
          <button id='tplQuickStats' label='Quick Stats' imageMso='FunctionsAutoSumInsertGallery' onAction='OnTemplate'
                  screentip='Quick Stats'
                  supertip='Count, sum, average, min and max for your selected range. Opens the editor seeded with your selection.'/>
          <button id='tplSumWhere' label='Sum Where' imageMso='AutoSum' onAction='OnTemplate'
                  screentip='Sum Where'
                  supertip='Conditional total over your selection (SUMIFS). Opens the editor seeded with your selection.'/>
          <button id='tplPivot' label='Pivot Generator' size='large' imageMso='PivotTableInsert' onAction='OnTemplate'
                  screentip='Pivot Generator'
                  supertip='Wizard: pick a group-by column and an aggregation to build a spilled summary, then review it in the editor.'/>
        </group>
        <group id='grpDates' label='Dates'>
          <button id='tplToday' label='Today' imageMso='DateAndTimePickerInsertContentControl' onAction='OnTemplate'
                  screentip='Today'
                  supertip=""Writes today's date (TODAY) straight into the selected cell.""/>
          <button id='tplDatePicker' label='Date Picker' imageMso='DateAndTimePickerInsertContentControl' onAction='OnTemplate'
                  screentip='Date Picker'
                  supertip='Pick a date on a calendar; it is written into the selected cell as a date.'/>
          <button id='tplAge' label='Date Between' imageMso='EquationInsert' onAction='OnTemplate'
                  screentip='Date Between'
                  supertip='Difference between the date to your left and today, into the selected cell. Opens the editor seeded with days/months/years variants (DATEDIF).'/>
        </group>
        <group id='grpHelp' label='Help'>
          <button id='btnDocs' label='Help &amp; Docs' size='large'
                  imageMso='Help' onAction='OnOpenDocs'
                  screentip='Open PExL documentation'
                  supertip='Open a searchable reference of every PExL tool and verb &#8212; each with a plain description, the Excel formula it compiles to, a live playground, and a &quot;Try it yourself!&quot; demo sheet.'/>
        </group>
      </tab>
    </tabs>
  </ribbon>
  <contextMenus>
    <contextMenu idMso='ContextMenuCell'>
      <button id='ctxOpenEditor' label='Edit with PExL' image='pexlLogo' onAction='OnOpenEditor'
              insertBeforeMso='Cut'/>
      <button id='ctxTranslate' label='Translate formula to PExL' imageMso='ConvertTextToTable' onAction='OnTranslate'
              insertBeforeMso='Cut'/>
      <menuSeparator id='ctxSep' insertBeforeMso='Cut'/>
    </contextMenu>
  </contextMenus>
</customUI>";
    }
}
