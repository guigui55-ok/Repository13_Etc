
C#（WinForms）のアプリで、↑キーと↓キーで、TextBoxの中身を切り替える機能を実装したいです。

対象の拡張子は、txtとiniにします。（これは、クラス内で定義して、コード上から変更できるようにします）

切り替え対象のデータは、リストで保持します。

データのセットについて、
TextBoxが空のときは、リストをリセット。
パスが入力されたら、そのディレクトリの対象の拡張子のファイルリストを取得して、ファイル名でソートします。

そして、現在入力されているファイル名を探して、ポジションをセットします。
ディレクトリが見つからないときは、リストを空にしてスキップします。

上記でセットが終わった後、上下キーを押すと、それに応じて現在のポジションを変更し、その要素の文字列をテキストボックスに反映します。
この時は、リスト更新をしません。

要素の最大時に↑キーを押すと、リストの最初に移動します。
要素が最小の時に↓キーを押すと、リストの最後に移動します。

データセット（リスト更新）のタイミングは、
テキストボックスが変更されたときとします。
（ファイル数が多くてBusyになることは許容します）

この機能は追加機能なので、完全に既存のクラスとは分離したいです。
なので、このクラスにFormを外部から受け取って、このクラス内でイベントを設置します。
---

リストについて、別クラスにしたほうが良いでしょうか？要素の最初と最後の動作が少し特殊なので。

---
まず設計方針を決めたいです。
不明点、疑問点、矛盾点、誤っている点、改善点などありましたら、助言をお願いします。


#####

方針としては、**リスト操作は別クラスに分けたほうが良い**です。

今回の機能は、ざっくり分けると以下の2つです。

```text
① TextBox / Form のイベント制御
② ファイルリストの保持・現在位置の移動
```

このうち、②はUIに依存しない純粋なロジックなので、別クラス化に向いています。

## 推奨構成

```text
FilePathTextBoxSwitcher
 ├─ TextBoxのTextChangedイベントを監視
 ├─ TextBoxのKeyDownイベントを監視
 ├─ ファイル一覧の再取得タイミングを制御
 └─ FilePathListNavigator を使って上下移動する

FilePathListNavigator
 ├─ List<string> を保持
 ├─ 現在位置を保持
 ├─ MoveNext()
 ├─ MovePrevious()
 ├─ SetItems(...)
 └─ 現在値を返す
```

## クラス分離したほうが良い理由

今回の仕様では、上下移動が少し特殊です。

> 最大位置で↑キー → 最初へ
> 最小位置で↓キー → 最後へ

通常の「前へ」「次へ」と少し違うため、TextBoxイベント内に直接書くと、後から読みにくくなります。

なので、例えば以下のように分けるとよいです。

```csharp
public class FilePathListNavigator
{
    private readonly List<string> _items = new List<string>();
    private int _position = -1;

    public bool HasItems
    {
        get { return _items.Count > 0; }
    }

    public string Current
    {
        get
        {
            if (_position < 0 || _position >= _items.Count)
            {
                return string.Empty;
            }

            return _items[_position];
        }
    }

    public void SetItems(List<string> items, string currentPath)
    {
        _items.Clear();
        _items.AddRange(items);

        if (_items.Count == 0)
        {
            _position = -1;
            return;
        }

        _position = _items.FindIndex(x =>
            string.Equals(x, currentPath, StringComparison.OrdinalIgnoreCase));

        if (_position < 0)
        {
            _position = 0;
        }
    }

    public string MoveUp()
    {
        if (_items.Count == 0)
        {
            return string.Empty;
        }

        if (_position >= _items.Count - 1)
        {
            _position = 0;
        }
        else
        {
            _position++;
        }

        return Current;
    }

    public string MoveDown()
    {
        if (_items.Count == 0)
        {
            return string.Empty;
        }

        if (_position <= 0)
        {
            _position = _items.Count - 1;
        }
        else
        {
            _position--;
        }

        return Current;
    }
}
```

※ ↑キーで次、↓キーで前、という仕様に見えますが、一般的には逆に感じる人もいます。
ここは一度決めておいたほうがよいです。

## 気になる点・確認したほうがよい点

一番気になるのは、**TextBoxのTextChangedでリスト更新するが、上下キーでTextBox.Textを書き換えると、またTextChangedが発生する**点です。

そのため、イベント内更新中はリスト更新しないように、フラグが必要です。

```csharp
private bool _isChangingByKey;
```

イメージはこうです。

```csharp
private void TextBox_KeyDown(object sender, KeyEventArgs e)
{
    if (e.KeyCode == Keys.Up)
    {
        _isChangingByKey = true;
        _textBox.Text = _navigator.MoveUp();
        _isChangingByKey = false;

        e.Handled = true;
        e.SuppressKeyPress = true;
    }
}

private void TextBox_TextChanged(object sender, EventArgs e)
{
    if (_isChangingByKey)
    {
        return;
    }

    // ここでリスト再取得
}
```

## 改善案

対象拡張子はクラス内定義でもよいですが、あとで変更しやすくするならプロパティにするとよいです。

```csharp
public string[] TargetExtensions { get; set; } = new[] { ".txt", ".ini" };
```

また、Formを受け取るより、基本は **TextBoxだけ受け取る** ほうが分離度は高いです。

```csharp
public FilePathTextBoxSwitcher(TextBox textBox)
```

Form全体を渡すと、この追加機能クラスがFormに依存しすぎるためです。

## 結論

おすすめはこの構成です。

```text
FilePathTextBoxSwitcher
→ UIイベント担当

FilePathListNavigator
→ リスト・現在位置・上下移動担当
```

特に `TextChanged` と `KeyDown` の相互発火を避けるため、**キー操作によるTextBox更新中はリスト更新しないフラグ**を入れるのが重要です。


#####

再度検討しました。
基本的な方針はそのままで、以下のようにします。

クラス構成は推奨の通りにします。
キー操作のリスト移動について、↑キーで要素位置をプラス、↓キーで要素位置をマイナスします。
クラス内に保持するのは、FormではなくTextBoxにします。
対象拡張子はプロパティにします。
TextBoxのTextChangedでリスト更新して、
リスト更新中をフラグで管理して、上下キーでTextBox.Textを書き換え時に、リスト更新を繰り返し実行しないようにする。

上記で特に問題なければ、実装に進みましょう。

#####

はい、その方針で問題ありません。
実装例は以下です。

## 1. FilePathListNavigator.cs

```csharp
using System;
using System.Collections.Generic;

namespace TaskStarter
{
    public class FilePathListNavigator
    {
        private readonly List<string> _items = new List<string>();
        private int _position = -1;

        public bool HasItems
        {
            get { return _items.Count > 0; }
        }

        public void Clear()
        {
            _items.Clear();
            _position = -1;
        }

        public void SetItems(List<string> items, string currentValue)
        {
            Clear();

            if (items == null || items.Count == 0)
            {
                return;
            }

            _items.AddRange(items);

            _position = _items.FindIndex(delegate (string x)
            {
                return string.Equals(x, currentValue, StringComparison.OrdinalIgnoreCase);
            });

            if (_position < 0)
            {
                _position = 0;
            }
        }

        public string MoveUp()
        {
            if (_items.Count == 0)
            {
                return string.Empty;
            }

            _position++;

            if (_position >= _items.Count)
            {
                _position = 0;
            }

            return _items[_position];
        }

        public string MoveDown()
        {
            if (_items.Count == 0)
            {
                return string.Empty;
            }

            _position--;

            if (_position < 0)
            {
                _position = _items.Count - 1;
            }

            return _items[_position];
        }
    }
}
```

## 2. FilePathTextBoxSwitcher.cs

```csharp
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace TaskStarter
{
    public class FilePathTextBoxSwitcher
    {
        private readonly TextBox _textBox;
        private readonly FilePathListNavigator _navigator;

        private bool _isUpdatingByKey;

        public string[] TargetExtensions { get; set; }

        public FilePathTextBoxSwitcher(TextBox textBox)
        {
            if (textBox == null)
            {
                throw new ArgumentNullException("textBox");
            }

            _textBox = textBox;
            _navigator = new FilePathListNavigator();

            TargetExtensions = new string[] { ".txt", ".ini" };

            _textBox.TextChanged += TextBox_TextChanged;
            _textBox.KeyDown += TextBox_KeyDown;
        }

        private void TextBox_TextChanged(object sender, EventArgs e)
        {
            if (_isUpdatingByKey)
            {
                return;
            }

            UpdateFileList();
        }

        private void TextBox_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Up)
            {
                SetTextByKey(_navigator.MoveUp());
                e.Handled = true;
                e.SuppressKeyPress = true;
            }
            else if (e.KeyCode == Keys.Down)
            {
                SetTextByKey(_navigator.MoveDown());
                e.Handled = true;
                e.SuppressKeyPress = true;
            }
        }

        private void SetTextByKey(string value)
        {
            if (string.IsNullOrEmpty(value))
            {
                return;
            }

            try
            {
                _isUpdatingByKey = true;

                _textBox.Text = value;
                _textBox.SelectionStart = _textBox.Text.Length;
            }
            finally
            {
                _isUpdatingByKey = false;
            }
        }

        private void UpdateFileList()
        {
            string currentPath = _textBox.Text;

            if (string.IsNullOrWhiteSpace(currentPath))
            {
                _navigator.Clear();
                return;
            }

            string directoryPath = Path.GetDirectoryName(currentPath);

            if (string.IsNullOrEmpty(directoryPath) || !Directory.Exists(directoryPath))
            {
                _navigator.Clear();
                return;
            }

            List<string> files = Directory.GetFiles(directoryPath)
                .Where(IsTargetExtension)
                .OrderBy(Path.GetFileName)
                .ToList();

            _navigator.SetItems(files, currentPath);
        }

        private bool IsTargetExtension(string filePath)
        {
            string extension = Path.GetExtension(filePath);

            foreach (string targetExtension in TargetExtensions)
            {
                if (string.Equals(extension, targetExtension, StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }
            }

            return false;
        }

        public void Detach()
        {
            _textBox.TextChanged -= TextBox_TextChanged;
            _textBox.KeyDown -= TextBox_KeyDown;
        }
    }
}
```

## 3. Form側の使用例

```csharp
private FilePathTextBoxSwitcher _filePathTextBoxSwitcher;

private void Form1_Load(object sender, EventArgs e)
{
    _filePathTextBoxSwitcher = new FilePathTextBoxSwitcher(this.FilePathTextBox);

    _filePathTextBoxSwitcher.TargetExtensions = new string[]
    {
        ".txt",
        ".ini"
    };
}
```

## 補足

この実装だと、例えば以下のように動きます。

```text
C:\Test\a.txt
C:\Test\b.ini
C:\Test\c.txt
```

`TextBox`に `C:\Test\b.ini` が入っている状態で、

```text
↑キー → C:\Test\c.txt
↓キー → C:\Test\a.txt
```

となります。

注意点として、`TextBox`にファイル名だけを入れる運用ではなく、**フルパスを入れる前提**です。
今回の仕様なら、この前提で進めるのが自然です。



