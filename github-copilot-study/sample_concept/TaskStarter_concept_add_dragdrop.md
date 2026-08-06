c#（Winforms）のについて、Formとその中のTextBoxにファイルをDrag&Dropした時に、そのTextBox.Textにファイルパスを表示する、という機能を既存のFormクラスに追加したいです。

機能追加のため、上記処理を別のクラスに分離したいです。

まず設計方針を決めたいです。
不明点、疑問点、矛盾点、誤っている点、改善点などありましたら、助言をお願いします。



####

方針としては、**Drag&Drop処理を専用クラスに分離する**で問題ないです。
既存Formに直接イベント処理を書くより、再利用しやすくなります。

おすすめは以下です。

```text
Form
 └ TextBoxFileDropHandler
      ├ DragEnter
      └ DragDrop
```

## 基本方針

専用クラスに、対象の `TextBox` を渡して初期化します。

```csharp
var dropHandler = new TextBoxFileDropHandler(this.FilePathTextBox);
```

クラス側で以下を行います。

```csharp
textBox.AllowDrop = true;
textBox.DragEnter += TextBox_DragEnter;
textBox.DragDrop += TextBox_DragDrop;
```

## Form側にもDropさせたい場合

「FormにファイルをDropしても、そのTextBoxに反映したい」なら、Formも対象に含めるとよいです。

```csharp
var dropHandler = new TextBoxFileDropHandler(this, this.FilePathTextBox);
```

この場合、クラスは以下を管理します。

```text
対象Control:
- Form
- TextBox

出力先:
- TextBox
```

## 気になる点・確認したい点

### 1. 複数ファイルDrop時の扱い

ここは決めた方がよいです。

候補は以下です。

```text
A. 先頭1件だけTextBox.Textに入れる
B. 複数ファイルは拒否する
C. 複数ファイルを改行区切りで入れる
```

今回の用途が「ファイルパス入力欄」なら、**B. 複数ファイルは拒否**か、**A. 先頭1件だけ採用**がよいです。

おすすめは **B** です。
誤操作に気づきやすいからです。

### 2. フォルダDropを許可するか

これも決めた方がよいです。

```text
ファイルのみ許可
フォルダのみ許可
ファイル・フォルダ両方許可
```

「ファイルパスを表示」とあるので、基本は **ファイルのみ許可**が自然です。

### 3. 拡張子チェックをするか

たとえば `.txt` や `.ini` だけ許可したい場合は、クラスにプロパティを持たせるとよいです。

```csharp
dropHandler.AllowedExtensions = new List<string> { ".txt", ".ini" };
```

将来的に使い回すなら、最初から入れておく価値があります。

### 4. TextBoxに直接反映するか、イベント通知にするか

シンプルにするなら、クラス内で直接これを実行します。

```csharp
_textBox.Text = filePath;
```

ただし、将来的に

```text
Dropされたファイルをログ出力したい
Drop後に別処理を呼びたい
TextBox以外にも反映したい
```

となるなら、イベントも用意するとよいです。

```csharp
public event Action<string> FileDropped;
```

おすすめは、まずは **TextBox直接反映 + イベント通知も可能** くらいが扱いやすいです。

## 改善案

クラス名は以下がよいと思います。

```csharp
TextBoxFileDropHandler
```

または、FormにもDropさせるなら少し汎用的に、

```csharp
FileDropToTextBoxHandler
```

こちらの方が意味が明確です。

## 推奨設計

```text
クラス名:
FileDropToTextBoxHandler

役割:
FormまたはTextBoxにファイルがDrag&Dropされたとき、
指定されたTextBox.Textにファイルパスを設定する。

主な仕様:
- Drop対象Controlを複数指定できる
- 出力先TextBoxを1つ持つ
- ファイルのみ許可
- 複数ファイルは拒否
- 拡張子制限はプロパティで設定可能
```

この方針で進めて問題ないと思います。


####

追加の検討をして、以下のようにしようと思います。

汎用的なクラスにして、他のアプリでも使用したいため、保持するメンバは List<Control>, List<Form>にしたいです。
複数ファイルDrop時は、A.先頭1件だけ採用します。
フォルダDropは拒否します。（ただ、他のアプリでの流用も考えているので、明示的にフォルダのときは、処理中断する実装は入れておきたいです）
拡張子チェックは不要です。
メンバにLoggerを持たせて、直接ログ出力します。

特に問題なければ、実装に進みましょう。


####

はい、この方針で実装に進めて問題ありません。
ただし、`Form` も `Control` なので、内部保持は **`List<Control>` に統一**するのがおすすめです。

以下の構成でよいです。

```text
FileDropToTextBoxHandler
- Drop対象Control一覧を保持
- Dropされた先頭1ファイルをTextBox.Textへ設定
- フォルダは拒否
- 拡張子チェックなし
- Loggerへ直接ログ出力
```

## 実装例

```csharp
using System;
using System.Collections.Generic;
using System.IO;
using System.Windows.Forms;

public class FileDropToTextBoxHandler : IDisposable
{
    private readonly List<Control> _dropControls = new List<Control>();
    private readonly TextBox _targetTextBox;
    private readonly IFileDropLogger _logger;

    public FileDropToTextBoxHandler(TextBox targetTextBox, IFileDropLogger logger)
    {
        if (targetTextBox == null)
        {
            throw new ArgumentNullException(nameof(targetTextBox));
        }

        _targetTextBox = targetTextBox;
        _logger = logger;
    }

    public void AddDropControl(Control control)
    {
        if (control == null)
        {
            return;
        }

        if (_dropControls.Contains(control))
        {
            return;
        }

        control.AllowDrop = true;
        control.DragEnter += Control_DragEnter;
        control.DragDrop += Control_DragDrop;

        _dropControls.Add(control);

        LogInfo("Drag&Drop対象Controlを追加しました。ControlName=" + control.Name);
    }

    private void Control_DragEnter(object sender, DragEventArgs e)
    {
        if (e.Data == null || !e.Data.GetDataPresent(DataFormats.FileDrop))
        {
            e.Effect = DragDropEffects.None;
            return;
        }

        e.Effect = DragDropEffects.Copy;
    }

    private void Control_DragDrop(object sender, DragEventArgs e)
    {
        try
        {
            if (e.Data == null || !e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                return;
            }

            string[] paths = e.Data.GetData(DataFormats.FileDrop) as string[];

            if (paths == null || paths.Length == 0)
            {
                return;
            }

            // 複数Drop時は先頭1件だけ採用
            string path = paths[0];

            if (Directory.Exists(path))
            {
                LogInfo("フォルダDropのため処理を中断しました。Path=" + path);
                return;
            }

            if (!File.Exists(path))
            {
                LogInfo("ファイルが存在しないため処理を中断しました。Path=" + path);
                return;
            }

            _targetTextBox.Text = path;

            LogInfo("ファイルパスをTextBoxへ設定しました。Path=" + path);
        }
        catch (Exception ex)
        {
            LogError(ex, "Drag&Drop処理でエラーが発生しました。");
        }
    }

    public void Dispose()
    {
        foreach (Control control in _dropControls)
        {
            control.DragEnter -= Control_DragEnter;
            control.DragDrop -= Control_DragDrop;
        }

        _dropControls.Clear();
    }

    private void LogInfo(string message)
    {
        if (_logger != null)
        {
            _logger.Info(message);
        }
    }

    private void LogError(Exception ex, string message)
    {
        if (_logger != null)
        {
            _logger.Error(ex, message);
        }
    }
}
```

## Logger用インターフェース

```csharp
using System;

public interface IFileDropLogger
{
    void Info(string message);
    void Error(Exception ex, string message);
}
```

## Form側の使用例

```csharp
private FileDropToTextBoxHandler _fileDropHandler;

private void Form1_Load(object sender, EventArgs e)
{
    IFileDropLogger logger = new AppLoggerAdapter(_appLogger);

    _fileDropHandler = new FileDropToTextBoxHandler(this.FilePathTextBox, logger);

    // FormにDropしてもTextBoxへ反映
    _fileDropHandler.AddDropControl(this);

    // TextBoxにDropしてもTextBoxへ反映
    _fileDropHandler.AddDropControl(this.FilePathTextBox);
}

protected override void OnFormClosed(FormClosedEventArgs e)
{
    if (_fileDropHandler != null)
    {
        _fileDropHandler.Dispose();
        _fileDropHandler = null;
    }

    base.OnFormClosed(e);
}
```

## 既存AppLoggerに合わせるAdapter例

```csharp
using System;

public class AppLoggerAdapter : IFileDropLogger
{
    private readonly AppLogger _logger;

    public AppLoggerAdapter(AppLogger logger)
    {
        _logger = logger;
    }

    public void Info(string message)
    {
        if (_logger != null)
        {
            _logger.info(message);
        }
    }

    public void Error(Exception ex, string message)
    {
        if (_logger != null)
        {
            _logger.err(ex, message);
        }
    }
}
```

結論として、`List<Form>` は不要で、`List<Control>` だけで十分です。FormもTextBoxもPanelも同じように登録できます。


