using System;
using IO = System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

using Microsoft.WindowsAPICodePack.Dialogs;
using Wsh = IWshRuntimeLibrary;
using System.IO;

namespace LnkToFile
{
    /// <summary>
    /// MainWindow.xaml の相互作用ロジック
    /// </summary>
    public partial class MainWindow : Window
    {

        public MainWindow()
        {
            InitializeComponent();
        }

        /// <summary>
        /// 参照ボタンクリック時のイベント
        /// </summary>
        /// <param name="sender">使用しない</param>
        /// <param name="e">使用しない</param>
        private void Button_Click_RefDir(object sender, RoutedEventArgs e)
        {
            var dialog = new CommonOpenFileDialog
            {
                IsFolderPicker = true,
                Title = "対象のフォルダを選択してください。",
                InitialDirectory = @"C:\"
            };

            if (dialog.ShowDialog() == CommonFileDialogResult.Ok)
            {
                this.TargetPathText.Text = dialog.FileName;
            }
        }

        /// <summary>
        /// 実行ボタンクリック時のイベント
        /// </summary>
        /// <param name="sender">使用しない</param>
        /// <param name="e">使用しない</param>
        private async void Button_Click_Enter(object sender, RoutedEventArgs e)
        {
            var targetPath = this.TargetPathText.Text;
            if (!IO.Directory.Exists(targetPath))
            {
                MessageBox.Show($"ディレクトリ「{targetPath}」が見つかりませんでした。", "エラー");
            }

            try
            {
                // 相互運用型のため、対応するインターフェースを通じてインスタンスを生成。
                var sh = new Wsh.WshShell();

                foreach (var file in IO.Directory.GetFiles(targetPath, "*.lnk", IO.SearchOption.AllDirectories))
                {
                    var filePath = ((Wsh.IWshShortcut)sh.CreateShortcut(file)).TargetPath;

                    if (IO.File.Exists(filePath))
                    {
                        // ボトルネック解消のため。
                        await Task.Run(() =>
                        {
                            var next = IO.Path.Combine(IO.Path.GetDirectoryName(file), IO.Path.GetFileName(filePath));
                            IO.File.Copy(filePath, next);

                            var fileInfo = new FileInfo(file);
                            if ((fileInfo.Attributes & FileAttributes.ReadOnly) == FileAttributes.ReadOnly)
                            {
                                fileInfo.Attributes = FileAttributes.Normal;
                            }

                            fileInfo.Delete();
                        });
                    }
                }

                MessageBox.Show("コピーが完了しました。", "完了");

            }
            catch (IOException)
            {
                MessageBox.Show("入出力に関する例外が発生しました。\n開発者にお問い合わせください。");
            }
            catch (OutOfMemoryException)
            {
                // 以下のcatchでOutOfMemoryExceptionをcatchしないため。
                throw;
            }
            catch
            {
                MessageBox.Show("不明な例外が発生しました。\n開発者にお問い合わせください。");
            }
        }
    }
}
