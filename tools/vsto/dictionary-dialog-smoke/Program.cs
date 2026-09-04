#nullable disable
using System;
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;

namespace ChuanHoa.DictionaryDialogSmoke
{
    internal static class Program
    {
        [STAThread]
        private static int Main()
        {
            try
            {
                Application.EnableVisualStyles();
                var assembly = Assembly.Load("ChuanHoa.AddIn.Vsto");
                var dialogType = assembly.GetType(
                    "ChuanHoa.AddIn.Vsto.Runtime.CustomDictionaryDialog", true);
                var prompt = dialogType.GetMethod("Prompt",
                    BindingFlags.Public | BindingFlags.Static);
                Assert(prompt != null, "Dictionary Prompt method is missing.");

                using (var dialog = (Form)Activator.CreateInstance(dialogType,
                    BindingFlags.Instance | BindingFlags.NonPublic, null,
                    new object[] { "word-session:dictionary-smoke", "Thuật ngữ được chọn" }, null))
                {
                    InspectDialog(dialog);
                }

                Console.WriteLine(
                    "DICTIONARY_DIALOG_SMOKE_PASS OWNER_CONTRACT DPI_100_125_150_200 EMPTY_ERROR_DUPLICATE_CONTROLS ENTER_ESCAPE_TAB_SELECTION_WORKFLOW");
                return 0;
            }
            catch (TargetInvocationException exception)
            {
                Console.Error.WriteLine("DICTIONARY_DIALOG_SMOKE_FAIL " +
                    (exception.InnerException ?? exception));
                return 1;
            }
            catch (Exception exception)
            {
                Console.Error.WriteLine("DICTIONARY_DIALOG_SMOKE_FAIL " + exception);
                return 1;
            }
        }

        private static void InspectDialog(Form dialog)
        {
            Assert(dialog.AutoScaleMode == AutoScaleMode.Dpi,
                "Dialog is not DPI-aware.");
            Assert(dialog.CancelButton != null,
                "Escape does not have a close action.");

            var search = Required<TextBox>(dialog, "txtDictionarySearch");
            var list = Required<ListBox>(dialog, "lstDictionaryWords");
            var delete = Required<Button>(dialog, "btnDeleteDictionaryWord");
            var clear = Required<Button>(dialog, "btnClearDocumentIgnores");
            var addSelection = Required<Button>(dialog, "btnAddSelectedText");
            var ignoreSelection = Required<Button>(dialog, "btnIgnoreSelectedText");
            var newWord = Required<TextBox>(dialog, "txtNewDictionaryWord");
            var add = Required<Button>(dialog, "btnAddDictionaryWord");
            var close = Required<Button>(dialog, "btnCloseDictionary");

            Assert(!delete.Enabled, "Delete must be disabled without a selection.");
            Assert(addSelection.Enabled && ignoreSelection.Enabled,
                "Selected/commented text workflow is not enabled for a document scope.");
            Assert(newWord.TabIndex < add.TabIndex && add.TabIndex < close.TabIndex,
                "New-word keyboard tab order is invalid.");
            Assert(search.TabIndex < list.TabIndex && list.TabIndex < delete.TabIndex &&
                delete.TabIndex < clear.TabIndex,
                "Dictionary navigation tab order is invalid.");
            Assert(!string.IsNullOrWhiteSpace(search.AccessibleName) &&
                !string.IsNullOrWhiteSpace(list.AccessibleName) &&
                !string.IsNullOrWhiteSpace(newWord.AccessibleName),
                "Text inputs/list are missing accessible names.");

            foreach (var scale in new[] { 1.0f, 1.25f, 1.5f, 2.0f })
            {
                using (var clone = (Form)Activator.CreateInstance(dialog.GetType(),
                    BindingFlags.Instance | BindingFlags.NonPublic, null,
                    new object[] { "word-session:dpi", "Thuật ngữ được chọn" }, null))
                {
                    clone.Scale(new SizeF(scale, scale));
                    foreach (Control control in clone.Controls)
                    {
                        if (!control.Visible) continue;
                        Assert(control.Right <= clone.ClientSize.Width + 1 &&
                            control.Bottom <= clone.ClientSize.Height + 1,
                            "Control is clipped at scale " + scale + ": " + control.Name);
                    }
                }
            }
        }

        private static T Required<T>(Control root, string name) where T : Control
        {
            var found = root.Controls.Find(name, true).OfType<T>().SingleOrDefault();
            Assert(found != null, "Required control is missing: " + name);
            return found;
        }

        private static void Assert(bool condition, string message)
        {
            if (!condition) throw new InvalidOperationException(message);
        }
    }
}
