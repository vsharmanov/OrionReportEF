using System;
using System.Collections;
using System.Drawing;
using System.Drawing.Printing;
using System.Windows.Forms;

namespace OrionApp
{
    internal class ClsPrint
    {
        #region Variables

        private readonly ArrayList _arrColumnLefts = new ArrayList();
        //Used to save left coordinates of columns
        private readonly ArrayList _arrColumnWidths = new ArrayList();

        private readonly DataGridView _gw;
        //Used to save column widths
        private readonly PrintDocument _printDocument = new PrintDocument();

        private readonly string _reportHeader;
        private bool _bFirstPage;
        //Used to check whether we are printing first page
        private bool _bNewPage;

        private int _iCellHeight; //Used to get/set the datagridview cell height
                                  // Used to check whether we are printing a new page
        private int _iHeaderHeight;

        private int _iRow;
        private int _iTotalWidth; //
                                  //Used as counter
                                  //Used for the header height
        private StringFormat _strFormat; //Used to format the grid rows.
        #endregion

        public ClsPrint(DataGridView gridview, string reportHeader)
        {
            _printDocument.PrintPage += _printDocument_PrintPage;
            _printDocument.BeginPrint += _printDocument_BeginPrint;
            _gw = gridview;
            _reportHeader = reportHeader;
        }

        public void PrintForm()
        {
            ////Open the print dialog
            //PrintDialog printDialog = new PrintDialog();
            //printDialog.Document = _printDocument;
            //printDialog.UseEXDialog = true;

            ////Get the document
            //if (DialogResult.OK == printDialog.ShowDialog())
            //{
            //    _printDocument.DocumentName = "Test Page Print";
            //    _printDocument.Print();
            //}

            //Open the print preview dialog
            var objPPdialog = new PrintPreviewDialog
            {
                Document = _printDocument
            };
            objPPdialog.ShowDialog();
        }

        private void _printDocument_BeginPrint(object sender, PrintEventArgs e)
        {
            try
            {
                _strFormat = new StringFormat
                {
                    Alignment = StringAlignment.Center,
                    LineAlignment = StringAlignment.Center,
                    Trimming = StringTrimming.EllipsisCharacter
                };

                _arrColumnLefts.Clear();
                _arrColumnWidths.Clear();
                _iCellHeight = 0;
                _iRow = 0;
                _bFirstPage = true;
                _bNewPage = true;

                // Calculating Total Widths
                _iTotalWidth = 0;
                foreach (DataGridViewColumn dgvGridCol in _gw.Columns) _iTotalWidth += dgvGridCol.Width;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void _printDocument_PrintPage(object sender, PrintPageEventArgs e)
        {
            //try
            //{
            //Set the left margin
            var iLeftMargin = e.MarginBounds.Left;
            //Set the top margin
            var iTopMargin = e.MarginBounds.Top;
            //Whether more pages have to print or not
            var bMorePagesToPrint = false;

            //For the first page to print set the cell width and header height
            if (_bFirstPage)
                foreach (DataGridViewColumn gridCol in _gw.Columns)
                {
                    var iTmpWidth = (int) Math.Floor(gridCol.Width /
                                                     (double) _iTotalWidth * _iTotalWidth *
                                                     (e.MarginBounds.Width / (double) _iTotalWidth));

                    _iHeaderHeight = (int) e.Graphics.MeasureString(gridCol.HeaderText,
                                         gridCol.InheritedStyle.Font, iTmpWidth).Height + 11;

                    // Save width and height of headers
                    _arrColumnLefts.Add(iLeftMargin);
                    _arrColumnWidths.Add(iTmpWidth);
                    iLeftMargin += iTmpWidth;
                }

            //Loop till all the grid rows not get printed
            while (_iRow <= _gw.Rows.Count - 1)
            {
                var gridRow = _gw.Rows[_iRow];
                //Set the cell height
                _iCellHeight = gridRow.Height + 5;
                var iCount = 0;
                //Check whether the current page settings allows more rows to print
                if (iTopMargin + _iCellHeight >= e.MarginBounds.Height + e.MarginBounds.Top)
                {
                    _bNewPage = true;
                    _bFirstPage = false;
                    bMorePagesToPrint = true;
                    break;
                }
                else
                {
                    if (_bNewPage)
                    {
                        //Draw Header
                        e.Graphics.DrawString(_reportHeader,
                            new Font(_gw.Font, FontStyle.Bold),
                            Brushes.Black, e.MarginBounds.Left,
                            e.MarginBounds.Top - e.Graphics.MeasureString(_reportHeader,
                                new Font(_gw.Font, FontStyle.Bold),
                                e.MarginBounds.Width).Height - 13);

                        var strDate = "";
                        //Draw Date
                        e.Graphics.DrawString(strDate,
                            new Font(_gw.Font, FontStyle.Bold), Brushes.Black,
                            e.MarginBounds.Left +
                            (e.MarginBounds.Width - e.Graphics.MeasureString(strDate,
                                 new Font(_gw.Font, FontStyle.Bold),
                                 e.MarginBounds.Width).Width),
                            e.MarginBounds.Top - e.Graphics.MeasureString(_reportHeader,
                                new Font(new Font(_gw.Font, FontStyle.Bold),
                                    FontStyle.Bold), e.MarginBounds.Width).Height - 13);

                        //Draw Columns                 
                        iTopMargin = e.MarginBounds.Top;
                        var gridCol = new DataGridViewColumn[_gw.Columns.Count];
                        var colcount = _gw.Columns.Count - 1;
                        //Convert ltr to rtl
                        foreach (DataGridViewColumn GridCol in _gw.Columns) gridCol[colcount--] = GridCol;
                        for (var i = gridCol.Length - 1; i >= 0; i--)
                        {
                            e.Graphics.FillRectangle(new SolidBrush(Color.LightGray),
                                new Rectangle((int) _arrColumnLefts[iCount], iTopMargin,
                                    (int) _arrColumnWidths[iCount], _iHeaderHeight));

                            e.Graphics.DrawRectangle(Pens.Black,
                                new Rectangle((int) _arrColumnLefts[iCount], iTopMargin,
                                    (int) _arrColumnWidths[iCount], _iHeaderHeight));

                            e.Graphics.DrawString(gridCol[i].HeaderText,
                                gridCol[i].InheritedStyle.Font,
                                new SolidBrush(gridCol[i].InheritedStyle.ForeColor),
                                new RectangleF((int) _arrColumnLefts[iCount], iTopMargin,
                                    (int) _arrColumnWidths[iCount], _iHeaderHeight), _strFormat);
                            iCount++;
                        }

                        _bNewPage = false;
                        iTopMargin += _iHeaderHeight;
                    }

                    iCount = 0;
                    var gridCell = new DataGridViewCell[gridRow.Cells.Count];
                    var cellcount = gridRow.Cells.Count - 1;
                    //Convert ltr to rtl
                    foreach (DataGridViewCell cel in gridRow.Cells) gridCell[cellcount--] = cel;
                    //Draw Columns Contents                
                    for (var i = gridCell.Length - 1; i >= 0; i--)
                    {
                        if (gridCell[i].Value != null)
                            e.Graphics.DrawString(gridCell[i].FormattedValue.ToString(),
                                gridCell[i].InheritedStyle.Font,
                                new SolidBrush(gridCell[i].InheritedStyle.ForeColor),
                                new RectangleF((int) _arrColumnLefts[iCount],
                                    iTopMargin,
                                    (int) _arrColumnWidths[iCount], _iCellHeight),
                                _strFormat);
                        //Drawing Cells Borders 
                        e.Graphics.DrawRectangle(Pens.Black,
                            new Rectangle((int) _arrColumnLefts[iCount], iTopMargin,
                                (int) _arrColumnWidths[iCount], _iCellHeight));
                        iCount++;
                    }
                }

                _iRow++;
                iTopMargin += _iCellHeight;
            }

            //If more lines exist, print another page.
            e.HasMorePages = bMorePagesToPrint;
            //}
            //catch (Exception exc)
            //{
            //    MessageBox.Show(exc.Message, "Error", MessageBoxButtons.OK,
            //       MessageBoxIcon.Error);
            //}
        }
    }
}