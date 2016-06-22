using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using ExcelDna.Integration;

namespace ExcelDNA.XlDialogs {
    /// <summary>
    ///     DIALOG.BOX(dialog_ref)
    ///     Dialog_ref    is a reference to a dialog box definition table on sheet, or an array containing the definition
    ///     table.
    /// </summary>
    public class XlDialog {
        /// <summary>
        ///     XlDialog 控件类型
        /// </summary>
        [Flags]
        public enum XlControl {
            /// <summary>
            ///     Default OK button	1
            /// </summary>
            XlDefaultOkButton = 1,

            /// <summary>
            ///     Cancel button	2
            /// </summary>
            XlCancelButton = 2,

            /// <summary>
            ///     OK button	3
            /// </summary>
            XlOkButton = 3,

            /// <summary>
            ///     Default Cancel button	4
            /// </summary>
            XlDefaultCancelButton = 4,

            /// <summary>
            ///     Static text	5
            /// </summary>
            XlStaticText = 5,

            /// <summary>
            ///     Text edit box	6
            /// </summary>
            XlTextBox = 6,

            /// <summary>
            ///     Integer edit box	7
            /// </summary>
            XlIntegerEedit = 7,

            /// <summary>
            ///     Number edit box	8
            /// </summary>
            XlNumberEdit = 8,

            /// <summary>
            ///     Formula edit box	9
            /// </summary>
            XlFormulaEdit = 9,

            /// <summary>
            ///     Reference edit box	10
            /// </summary>
            XlReferenceEdit = 10,

            /// <summary>
            ///     Option button group	11
            /// </summary>
            XlOptionButtonGgroup = 11,

            /// <summary>
            ///     Option button	12
            /// </summary>
            XlOptionButton = 12,

            /// <summary>
            ///     Check box	13
            /// </summary>
            XlCheckBox = 13,

            /// <summary>
            ///     Group box	14
            /// </summary>
            XlGroupBox = 14,

            /// <summary>
            ///     List box	15
            /// </summary>
            XlListbox = 15,

            /// <summary>
            ///     Linked list box	16
            /// </summary>
            XlLinkedListbox = 16,

            /// <summary>
            ///     Icons	17
            /// </summary>
            XlIcons = 17,
            //Linked file list box (Microsoft Excel for Windows only)	18
            //Linked drive and directory box (Microsoft Excel for Windows only)	19
            XlDirectoryTextbox = 20,

            /// <summary>
            ///     Drop-down list box	21
            /// </summary>
            XlDropdownList = 21,

            /// <summary>
            ///     Drop-down combination edit/list box	22
            /// </summary>
            XlCombobox = 22,

            /// <summary>
            ///     Picture button	23
            /// </summary>
            XlPictureButton = 23,

            /// <summary>
            ///     Help button	24
            /// </summary>
            XlHelpButton = 24,

            /// <summary>
            /// disable + ItemNum
            /// </summary>
            XlDisable = 200,

            /// <summary>
            ///     空类型
            /// </summary>
            XlEmpty = -1
        }

        /// <summary>
        ///     Value 项索引
        /// </summary>
        private const int ControlValueIndex = 6;

        /// <summary>
        ///     Text项索引
        /// </summary>
        protected const int ItemIndexText = 5;

        public readonly XlDialogControlCollections Controls = new XlDialogControlCollections();

        /// <summary>
        ///     结果数组
        /// </summary>
        private object[,] _resultValue;

        /// <summary>
        /// 窗体定义
        /// </summary>
        /// <remarks>
        ///     The first row of dialog_ref defines the position, size, and name of the dialog box.
        ///     It can also specify the default selected item and the reference for the Help button.
        ///     The position is specified in columns 2 and 3, the size in columns 4 and 5, and the name in column 6.
        ///     To specify a default item, place the item's position number in column 7.
        ///     You can place the reference for the Help button in row 1, column 1 of the table,
        ///     but the preferred location is column 7 in the row where the Help button is defined. Row 1, column 1 is usually left
        ///     blank.
        /// </remarks>
        private readonly ControlItem _formControl;

        public XlDialog() {
            _formControl = new ControlItem(XlControl.XlEmpty);
            this.Controls.Add(_formControl);
            Width = 300;
            Height = 200;
            Text = "Text";
        }

        /// <summary>
        ///     对话框位置 X
        /// </summary>
        public int X {
            get { return _formControl.X; }
            set { _formControl.X = value; }
        }

        /// <summary>
        ///     对话框 位置 Y
        /// </summary>
        public int Y {
            get { return _formControl.Y; }
            set { _formControl.Y = value; }
        }

        /// <summary>
        ///     对话框宽度
        /// </summary>
        public int Width {
            get { return _formControl.Width; }
            set { _formControl.Width = value; }
        }

        /// <summary>
        ///     对话框高度
        /// </summary>
        public int Height {
            get { return _formControl.Height; }
            set { _formControl.Height = value; }
        }

        /// <summary>
        ///     对话框标题
        /// </summary>
        public string Text {
            get { return _formControl.Text; }
            set { _formControl.Text = value; }
        }

        /// <summary>
        ///     显示对话框
        /// </summary>
        /// <returns></returns>
        public virtual bool ShowDialog() {
            try {
                var dialogDef = Controls.Build();
                var result = XlCall.Excel(XlCall.xlfDialogBox, dialogDef);
                _resultValue = result as object[,];
                if (_resultValue != null) {
                    Controls.UpdateResult(_resultValue);
                    return true;
                }
                return false;
            } finally {
                this.Controls.Dispose();
            }
        }

        /// <summary>
        ///     XlDialog 控件接口
        /// </summary>
        private interface IXlDialogControl : IDisposable {
            /// <summary>
            ///     控件类型
            /// </summary>
            XlControl ItemNumber { get; }

            /// <summary>
            ///     X 坐标，如果数值小于 0 则表示使用默认值
            /// </summary>
            int X { get; set; }

            /// <summary>
            ///     Y 坐标，如果数值小于 0 则表示使用默认值
            /// </summary>
            int Y { get; set; }

            /// <summary>
            ///     宽度,如果数值小于 0 则表示使用默认值
            /// </summary>
            int Width { get; set; }

            /// <summary>
            ///     高度，如果数值小于 0 则表示使用默认值
            /// </summary>
            int Height { get; set; }

            /// <summary>
            ///     文本内容
            /// </summary>
            string Text { get; set; }
        }

        #region Control Types
        /// <summary>
        ///     控件基类
        /// </summary>
        public class ControlItem : IXlDialogControl {
            public const int DefaultHeight = 20;
            /// <summary>
            /// 控件定义数组
            /// </summary>
            protected readonly object[] ControlDefine = new object[7];

            protected internal ControlItem(XlControl itemNumber) {
                ItemNumber = itemNumber;
                Visible = true;
            }

            /// <summary>
            ///     控件索引
            /// </summary>
            internal int Index { get; set; }

            internal object this[int index] {
                get { return ControlDefine[index]; }
                set { ControlDefine[index] = value; }
            }

            /// <summary>
            ///     控件类型
            /// </summary>
            public XlControl ItemNumber {
                get {
                    if (ControlDefine[0] == null) {
                        return XlControl.XlEmpty;
                    }
                    return (XlControl)ControlDefine[0];
                }
                protected set {
                    if (value < 0) {
                        ControlDefine[0] = null;
                    } else {
                        ControlDefine[0] = (int)value;
                    }
                }
            }

            /// <summary>
            ///     X 坐标，如果数值小于 0 则表示使用默认值
            /// </summary>
            public virtual int X {
                get {
                    if (ControlDefine[1] == null) {
                        return -1;
                    }
                    return (int)ControlDefine[1];
                }
                set {
                    if (value < 0) {
                        ControlDefine[1] = null;
                    } else {
                        ControlDefine[1] = value;
                    }
                }
            }

            /// <summary>
            ///     Y 坐标，如果数值小于 0 则表示使用默认值
            /// </summary>
            public virtual int Y {
                get {
                    if (ControlDefine[2].IsNull()) {
                        return -1;
                    }
                    return (int)ControlDefine[2];
                }
                set {
                    if (value < 0) {
                        ControlDefine[2] = null;
                    } else {
                        ControlDefine[2] = value;
                    }
                }
            }

            /// <summary>
            ///     宽度,如果数值小于 0 则表示使用默认值
            /// </summary>
            public virtual int Width {
                get {
                    if (ControlDefine[3].IsNull()) {
                        return -1;
                    }
                    return (int)ControlDefine[3];
                }
                set {
                    if (value < 0) {
                        ControlDefine[3] = null;
                    } else {
                        ControlDefine[3] = value;
                    }
                }
            }

            /// <summary>
            ///     高度，如果数值小于 0 则表示使用默认值
            /// </summary>
            public virtual int Height {
                get {
                    if (ControlDefine[4].IsNull()) {
                        return -1;
                    }
                    return (int)ControlDefine[4];
                }
                set {
                    if (value < 0) {
                        ControlDefine[4] = -1;
                    } else {
                        ControlDefine[4] = value;
                    }
                }
            }

            /// <summary>
            ///     文本内容
            /// </summary>
            public virtual string Text {
                get { return (string)ControlDefine[5]; }
                set { ControlDefine[5] = value; }
            }

            /// <summary>
            /// 控件是否可用
            /// </summary>
            public bool Enable {
                get {
                    if (ItemNumber == XlControl.XlEmpty) {
                        //窗体定义
                        return true;
                    }
                    return ItemNumber < XlControl.XlDisable;
                }
                set {
                    if (ItemNumber != XlControl.XlEmpty) {
                        if (value != Enable) {
                            if (value) {
                                ItemNumber -= XlControl.XlDisable;
                            } else {
                                ItemNumber += (int)XlControl.XlDisable;
                            }
                        }
                    }
                }
            }

            /// <summary>
            /// 控件是否可见
            /// </summary>
            [DefaultValue(true)]
            public bool Visible { get; set; }

            /// <summary>
            /// 活动控件定义数组
            /// </summary>
            /// <returns></returns>
            public virtual IEnumerable<object[]> GetControlDefine() {
                return new[] { ControlDefine };
            }

            #region Implementation of IDisposable

            /// <summary>
            ///     执行与释放或重置非托管资源相关的应用程序定义的任务。
            /// </summary>
            public virtual void Dispose() {
            }

            #endregion

            /// <summary>
            ///     构建对话框之前调用
            /// </summary>
            protected internal virtual void OnBeforeBuild() {
            }
        }

        /// <summary>
        ///     单元格引用 编辑器
        /// </summary>
        public class ReferenceEdit : ControlItem {
            public ReferenceEdit() : base(XlControl.XlReferenceEdit) {
            }

/*            /// <summary>
            ///     引用单元格地址
            /// </summary>
            public CellAddress Address {
                get {
                    if (Value.IsNull()) {
                        return null;
                    }
                    var sheetName = (string)XlCall.Excel(XlCall.xlfGetDocument, 76);
                    return $"{sheetName}!{Value}";
                }
            }*/

            /// <summary>
            ///     Reference 地址 (R1C1)
            /// </summary>
            public string Value {
                get { return (string)ControlDefine[ControlValueIndex]; }
                set { ControlDefine[ControlValueIndex] = value; }
            }
        }

        /// <summary>
        ///     公式编辑器
        /// </summary>
        public class FormulaEdit : ControlItem {
            public FormulaEdit() : base(XlControl.XlFormulaEdit) {
            }

            /// <summary>
            ///     公式内容
            /// </summary>
            public string Formula {
                get { return (string)ControlDefine[ControlValueIndex]; }
                set { ControlDefine[ControlValueIndex] = value; }
            }
        }

        /// <summary>
        ///     静态文本空间
        /// </summary>
        public sealed class Label : ControlItem {
            public Label() : base(XlControl.XlStaticText) {
            }

            public Label(string text) : this() {
                Text = text;
            }
        }

        /// <summary>
        ///     Group 框
        /// </summary>
        public sealed class GroupBox : ControlItem {
            public GroupBox() : base(XlControl.XlGroupBox) {
                X = Y = 10;
            }

            public GroupBox(string text) : this() {
                Text = text;
            }
        }

        /// <summary>
        ///     文本编辑控件
        /// </summary>
        public class TextBox : ControlItem {
            public TextBox() : base(XlControl.XlTextBox) {
            }

            public TextBox(string text) : this() {
                Value = text;
            }

            /// <summary>
            ///     文本框编辑内容
            /// </summary>
            public string Value {
                get { return (string)ControlDefine[ControlValueIndex]; }
                set { ControlDefine[ControlValueIndex] = value; }
            }
        }

        #region button Control
        /// <summary>
        ///     确定 按钮
        /// </summary>
        public sealed class OkButton : ControlItem {
            public OkButton() : base(XlControl.XlOkButton) {
                Text = "OK";
            }

            public OkButton(string text) : base(XlControl.XlOkButton) {
                Text = text;
            }

            /// <summary>
            /// 是否为默认按钮
            /// </summary>
            public bool Default {
                get { return ItemNumber == XlControl.XlDefaultOkButton; }
                set { ItemNumber = value ? XlControl.XlDefaultOkButton : XlControl.XlOkButton; }
            }
        }

        /// <summary>
        ///     取消按钮
        /// </summary>
        public sealed class CancelButton : ControlItem {
            public CancelButton() : base(XlControl.XlCancelButton) {
                Text = "Cancel";
            }

            /// <summary>
            /// 是否为默认按钮
            /// </summary>
            public bool Default {
                get { return ItemNumber == XlControl.XlDefaultCancelButton; }
                set { ItemNumber = value ? XlControl.XlDefaultCancelButton : XlControl.XlCancelButton; }
            }
        }

        #endregion Buttons

        #region List Box Controls
        public abstract class AbstractListControl:ControlItem {

            protected AbstractListControl(XlControl itemNumber) : base(itemNumber) {
                this.Items = new StringCollection();
            }

            /// <summary>
            /// 选中列表从 0 开始的索引
            /// 没有选中为 -1
            /// </summary>
            /// <remarks>内置的 索引从 1 开始，外部表现为复合.Net 一般约定的 从 0 开始 的索引 </remarks>
            [DefaultValue(-1)]
            public int SelectedIndex {
                get {
                    if (ControlDefine[ControlValueIndex].IsNull()) {
                        return -1;
                    }
                    return  Convert.ToInt32(ControlDefine[ControlValueIndex]) - 1;
                }
                set {
                    if (value < 0) {
                        ControlDefine[ControlValueIndex] = null;
                    } else {
                        ControlDefine[ControlValueIndex] = value + 1;
                    }
                }
            }

            /// <summary>
            ///     获取或者设置一个对象，该对象表示此 <see cref="T:ComboBox"/> 中所含的项的集合
            /// </summary>
            public StringCollection Items { get; }

            /// <summary>
            /// 列表名称
            /// </summary>
            private string ListName {
                get { return base.Text; }
                set { base.Text = value; }
            }

            /// <summary>
            ///     构建对话框之前调用
            /// </summary>
            protected internal override void OnBeforeBuild() {
                string[] listArray;
                if (Items != null && Items.Any()) {
                    listArray = Items.ToArray();
                } else {
                    //必须有一个列表否则会发生错误
                    listArray = new[] { string.Empty };
                }
                if (string.IsNullOrEmpty(ListName)) {
                    ListName = $"Gen_{GetType().Name}_{Index}";
                }
                if (SelectedIndex > Items.Count) {
                    SelectedIndex = -1;
                }
                XlCall.Excel(XlCall.xlfSetName, ListName, listArray);
            }

            #region Overrides of ControlItem

            /// <summary>
            ///     执行与释放或重置非托管资源相关的应用程序定义的任务。
            /// </summary>
            public override void Dispose() {
                if ((bool)XlCall.Excel(XlCall.xlfSetName, ListName)) {
                    XlCall.Excel(XlCall.xlFree, ListName);
                };
                base.Dispose();
            }
            #endregion

            public class StringCollection : Collection<String> {

                public void AddRange(IEnumerable<string> items) {
                    foreach (var item in items) {
                        this.Add(item);
                    }
                }
            }
        }

        /// <summary>
        ///     组合框 控件
        /// </summary>
        /// <remarks>
        ///     组合框之前需要 一个编辑项
        /// </remarks>
        public class ComboBox : AbstractListControl {
            private readonly TextBox _innerTextBox = new TextBox();

            /// <summary>
            /// </summary>
            public ComboBox() : base(XlControl.XlCombobox) {
            }

            /// <summary>
            ///     X 坐标，如果数值小于 0 则表示使用默认值
            /// </summary>
            public override int X {
                get { return base.X; }
                set {
                    base.X = value;
                    _innerTextBox.X = value;
                }
            }


            /// <summary>
            ///     Y 坐标，如果数值小于 0 则表示使用默认值
            /// </summary>
            public override int Y {
                get { return base.Y; }
                set {
                    base.Y = value;
                    _innerTextBox.Y = value;
                }
            }

            /// <summary>
            ///     宽度,如果数值小于 0 则表示使用默认值
            /// </summary>
            public override int Width {
                get { return base.Width; }
                set {
                    base.Width = value;
                    _innerTextBox.Width = value;
                }
            }

            /// <summary>
            ///     高度，如果数值小于 0 则表示使用默认值
            /// </summary>
            public override int Height {
                get { return base.Height; }
                set {
                    base.Height = value;
                    _innerTextBox.Height = value;
                }
            }

            /// <summary>
            ///     文本内容
            /// </summary>
            public override string Text {
                get { return _innerTextBox.Value; }
                set {
                    _innerTextBox.Value = value;
                    if (this.Items != null) {
                        SelectedIndex = Items.IndexOf(value);
                    }
                }
            }

            #region Overrides of ControlItem

            /// <summary>
            /// 活动控件定义数组
            /// </summary>
            /// <returns></returns>
            public override IEnumerable<object[]> GetControlDefine() {
                return new[] { _innerTextBox.GetControlDefine().FirstOrDefault(), base.ControlDefine };
            }

            #endregion
        }

        /// <summary>
        ///     下拉列表控件
        /// </summary>
        public class DropdownList : AbstractListControl {
            public DropdownList() : base(XlControl.XlDropdownList) {
            }

            /// <summary>
            ///     选定的值
            /// </summary>
            public string Value {
                get {
                    var index = SelectedIndex;
                    if (index < 0) {
                        return string.Empty;
                    }
                    return Items.ElementAt(index);
                }
            }
        }

        public class Listbox : AbstractListControl {
            public Listbox() : base(XlControl.XlListbox) {
            }
        }
        #endregion List box

        public class IntegerEedit : ControlItem {
            public IntegerEedit() : base(XlControl.XlIntegerEedit) {
            }

            public int Value {
                get { return Convert.ToInt32(ControlDefine[ControlValueIndex]); }
                set { ControlDefine[ControlValueIndex] = value; }
            }
        }

        public class NumberEdit : ControlItem {
            public NumberEdit() : base(XlControl.XlNumberEdit) {
            }

            public double Value {
                get { return Convert.ToDouble(ControlDefine[ControlValueIndex]); }
                set { ControlDefine[ControlValueIndex] = value; }
            }
        }

        public sealed class CheckBox : ControlItem {
            public CheckBox() : base(XlControl.XlCheckBox) {
            }

            public CheckBox(string text) : base(XlControl.XlCheckBox) {
                this.Text = text;
            }

            public bool Value {
                get {
                    if ( ControlDefine[ControlValueIndex].IsNull()) {
                        return false;
                    }
                    return (bool)ControlDefine[ControlValueIndex];
                }
                set { ControlDefine[ControlValueIndex] = value; }
            }
        }

        #endregion Control Types
        /// <summary>
        ///     对话框控件集合
        /// </summary>
        public class XlDialogControlCollections : Collection<ControlItem>, IDisposable {
            internal XlDialogControlCollections() {
            }

            #region Implementation of IDisposable

            /// <summary>
            ///     执行与释放或重置非托管资源相关的应用程序定义的任务。
            /// </summary>
            public void Dispose() {
                foreach (var item in Items) {
                    item.Dispose();
                }
            }

            #endregion

            /// <summary>
            ///     更新元素索引
            /// </summary>
            private void UpdateItemsIndex(int startIndex) {
                for (var i = startIndex; i < Count; i++) {
                    Items[i].Index = i;
                }
            }

            #region Overrides of Collection<ControlItem>

            /// <summary>
            ///     将元素插入 <see cref="T:System.Collections.ObjectModel.Collection`1" /> 的指定索引处。
            /// </summary>
            /// <param name="index">从零开始的索引，应在该位置插入 <paramref name="item" />。</param>
            /// <param name="item">要插入的对象。对于引用类型，该值可以为 null。</param>
            /// <exception cref="T:System.ArgumentOutOfRangeException">
            ///     <paramref name="index" /> 小于零。- 或 -<paramref name="index" /> 大于
            ///     <see cref="P:System.Collections.ObjectModel.Collection`1.Count" />。
            /// </exception>
            protected override void InsertItem(int index, ControlItem item) {
                base.InsertItem(index, item);
                item.Index = index;
            }

            /// <summary>
            ///     移除 <see cref="T:System.Collections.ObjectModel.Collection`1" /> 的指定索引处的元素。
            /// </summary>
            /// <param name="index">要移除的元素的从零开始的索引。</param>
            /// <exception cref="T:System.ArgumentOutOfRangeException">
            ///     <paramref name="index" /> 小于零。- 或 -<paramref name="index" />
            ///     等于或大于 <see cref="P:System.Collections.ObjectModel.Collection`1.Count" />。
            /// </exception>
            protected override void RemoveItem(int index) {
                base.RemoveItem(index);
                UpdateItemsIndex(index);
            }

            /// <summary>
            /// 构建 控件定义数组
            /// </summary>
            /// <returns></returns>
            internal object[,] Build() {
                var visibleControls = Items.Where(i => i.Visible).ToArray();
                var rows = visibleControls.Sum(i => i.GetControlDefine().Count());
                var result = new object[rows, 7];
                int rowIndex = 0;
                foreach (var item in visibleControls) {
                    item.OnBeforeBuild();
                    var defArray = item.GetControlDefine();
                    foreach (var array in defArray) {
                        for (int i = 0; i < 7; i++) {
                            result[rowIndex, i] = array[i];
                        }
                        rowIndex++;
                    }
                }
                return result;
            }

            /// <summary>
            ///     解析返回值,会写到控件集合
            /// </summary>
            internal void UpdateResult(object[,] result) {
                int index = 0;
                var visibleControls = Items.Where(i => i.Visible).ToArray();
                foreach (var item in visibleControls) {
                    var controlDefs = item.GetControlDefine();
                    foreach (var defItem in controlDefs) {
                        defItem[ControlValueIndex] = result[index, ControlValueIndex];
                        index++;
                    }
                }
            }

            #endregion
        }
    }
}