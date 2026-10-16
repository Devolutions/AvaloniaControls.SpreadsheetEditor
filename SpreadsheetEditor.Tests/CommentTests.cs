using System.IO;
using System.Linq;
using System.Threading.Tasks;

using Avalonia.Controls;
using Avalonia.Headless.XUnit;
using Avalonia.Interactivity;
using Avalonia.VisualTree;

using ClosedXML.Excel;

using SpreadsheetEditor.Controls;
using SpreadsheetEditor.Controls.Localization;
using SpreadsheetEditor.Controls.Models;
using SpreadsheetEditor.Controls.Xlsx;
using EditorControl = SpreadsheetEditor.Controls.SpreadsheetEditor;

using Xunit;

namespace SpreadsheetEditor.Tests;

public sealed class CommentTests
{
    [AvaloniaFact]
    public async Task CreatingCommentDoesNotAddAnAuthorLabelAndSurvivesSave()
    {
        EditorControl editor = new();
        editor.NewSheet(1, 1);
        Window window = new() { Content = editor, Width = 500, Height = 300 };
        window.Show();
        try
        {
            window.UpdateLayout();
            SpreadsheetGrid grid = editor.FindControl<SpreadsheetGrid>("Grid")!;
            MenuFlyout menu = Assert.Single(grid.GetVisualDescendants().OfType<Control>()
                .Select(control => control.ContextFlyout).OfType<MenuFlyout>(),
                flyout => flyout.Items.OfType<MenuItem>()
                    .Any(item => Equals(item.Header, EditorText.Get(EditorString.InsertComment))));
            MenuItem insert = Assert.Single(menu.Items.OfType<MenuItem>(),
                item => Equals(item.Header, EditorText.Get(EditorString.InsertComment)));
            insert.RaiseEvent(new RoutedEventArgs(MenuItem.ClickEvent));
            Window dialog = Assert.Single(window.OwnedWindows);
            Assert.Single(dialog.GetVisualDescendants().OfType<TextBox>()).Text = "My comment";
            Button save = Assert.Single(dialog.GetVisualDescendants().OfType<Button>(), button => button.IsDefault);
            save.RaiseEvent(new RoutedEventArgs(Button.ClickEvent));
            // Dialog continuations apply the comment on the dispatcher.
            await Task.Yield();
            Assert.Equal("My comment", editor.ViewModel.Rows[0].Cells[0].Comment);
            Assert.Null(editor.ViewModel.Rows[0].Cells[0].CommentAuthor);

            using MemoryStream saved = new();
            editor.SaveXlsx(saved);
            saved.Position = 0;
            using (XLWorkbook workbook = new(saved))
            {
                IXLComment comment = workbook.Worksheet(1).Cell(1, 1).GetComment();
                Assert.Equal("My comment", comment.Text);
                Assert.Equal(string.Empty, comment.Author);
            }
            saved.Position = 0;
            CellData reopened = XlsxFile.Read(saved).Cells[0][0];
            Assert.Equal("My comment", reopened.Comment);
            Assert.True(string.IsNullOrEmpty(reopened.CommentAuthor));
        }
        finally
        {
            window.Close();
        }
    }

    [AvaloniaTheory]
    [InlineData(null)]
    [InlineData("Spreadsheet Editor")]
    [InlineData("Original author")]
    public void CommentTooltipContainsOnlyCommentTextWhilePreservingAuthorMetadata(string? author)
    {
        EditorControl editor = new();
        editor.LoadSheet(new SheetData
        {
            Cells = new[] { new[] { CellData.Plain("Value") with { Comment = "My comment", CommentAuthor = author } } },
        });
        Window window = new() { Content = editor, Width = 500, Height = 300 };
        window.Show();
        try
        {
            window.UpdateLayout();
            SpreadsheetGrid grid = editor.FindControl<SpreadsheetGrid>("Grid")!;
            Assert.Single(grid.GetVisualDescendants().OfType<Border>(),
                border => Equals(ToolTip.GetTip(border), "My comment"));
            Assert.Equal(author, editor.GetCellData()[0][0].CommentAuthor);
        }
        finally
        {
            window.Close();
        }
    }
}
