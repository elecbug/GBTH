using GBTH.List;

namespace GBTH
{
    public partial class MainForm : Form
    {
        private SplitContainer tool_container;  // 최상위 수평 분할 컨테이너
        private SplitContainer container;       // 하단 수직 분할 컨테이너
        private IngredientList list_view;       // 왼쪽 리스트
        private ReportGrid grid_view;           // 오른쪽 표
        private NumericUpDown numeric;          // 년도 선택 컨트롤
        private Button print;                   // 엑셀 변환 버튼
        private Label label;                    // 변환 퍼센트 레이블
        private Button sync;                    // 동기화 버튼
        private ContextMenuStrip strip;         // 리스트 뷰 우클릭 다이얼로그

        public MainForm()
        {
            this.Load += MainFormLoad;
            this.FormClosed += MainFormFormClosed;

            InitializeComponent();

            this.Text = "GBTH";
            this.WindowState = FormWindowState.Maximized;
            this.Font = new Font("굴림체", 13, FontStyle.Regular);
            this.Icon = Properties.Resources.Katty;

            this.tool_container = new SplitContainer()
            {
                Parent = this,
                Visible = true,
                Size = this.ClientSize,
                Orientation = Orientation.Horizontal,
                Dock = DockStyle.Fill,
            };

            this.container = new SplitContainer()
            {
                Parent = this.tool_container.Panel2,
                Visible = true,
                Size = this.tool_container.Panel1.ClientSize,
                Dock = DockStyle.Fill,
            };

            this.grid_view = new ReportGrid()
            {
                Parent = this.container.Panel2,
                Visible = true,
                Size = this.container.Panel2.ClientSize,
                Dock = DockStyle.Fill,
            };

            this.numeric = new NumericUpDown()
            {
                Location = new Point(5, 5),
                Parent = this.tool_container.Panel1,
                Visible = true,
                Minimum = 2000,
                Maximum = 3000,
                Value = DateTime.Now.Year,
                Increment = 1,
            };
            this.numeric.ValueChanged += NumericValueChanged;

            this.print = new Button()
            {
                Location = new Point(this.numeric.Width + 10, 5),
                Parent = this.tool_container.Panel1,
                Size = this.numeric.Size,
                Visible = true,
                Text = "엑셀 생성",
            };
            this.print.Click += PrintClick;

            this.label = new Label()
            {
                Location = new Point(this.numeric.Width + 10, this.print.Height + 5),
                Parent = this.tool_container.Panel1,
                Size = this.numeric.Size,
                Visible = true,
                Text = "",
            };

            this.sync = new Button()
            {
                Location = new Point(this.numeric.Width * 2 + 15, 5),
                Parent = this.tool_container.Panel1,
                Size = this.numeric.Size,
                Visible = true,
                Text = "동기화",
            };
            this.sync.Click += SyncClick;

            this.list_view = IngredientList.Deserialize(ref this.grid_view, Environment.CurrentDirectory, (int)this.numeric.Value);
            this.list_view.Parent = this.container.Panel1;
            this.list_view.Visible = true;
            this.list_view.Size = this.container.Panel1.ClientSize;
            this.list_view.Dock = DockStyle.Fill;
            this.list_view.MouseClick += ListViewMouseClick;

            this.strip = new ContextMenuStrip();
            this.strip.Items.Add("수정").Click += (sender, e) => { (sender as IngredientList)!.Edit(); };
            this.strip.Items.Add("추가").Click += (sender, e) => { (sender as IngredientList)!.Add(); };
            this.strip.Items.Add("삭제").Click += (sender, e) => { (sender as IngredientList)!.Remove(); };
        }

        private void ListViewMouseClick(object? sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {

                this.strip.Show(Control.MousePosition.X, Control.MousePosition.Y);
            }
        }

        private void MainFormLoad(object? sender, EventArgs e)
        {
            PropertiesLoad();
        }
        private void MainFormFormClosed(object? sender, FormClosedEventArgs e)
        {
            this.list_view.SaveData();

            PropertiesSave();
        }

        private void NumericValueChanged(object? sender, EventArgs e)
        {
            PropertiesSave();

            this.list_view.Visible = false;
            this.list_view.Dispose();

            this.list_view = IngredientList.Deserialize(ref this.grid_view, Environment.CurrentDirectory, (int)this.numeric.Value);
            this.list_view.Parent = this.container.Panel1;
            this.list_view.Visible = true;
            this.list_view.Size = this.container.Panel1.ClientSize;
            this.list_view.Dock = DockStyle.Fill;

            PropertiesLoad();
        }
        private void PrintClick(object? sender, EventArgs e)
        {
            this.Invoke((MethodInvoker)delegate
            { 
                this.list_view.Print(this.label); 
            });
        }

        private void SyncClick(object? sender, EventArgs e)
        {

        }

        private void PropertiesSave()
        {
            Properties.Settings.Default.AllSettings
                = this.tool_container.SplitterDistance + "\\"
                + this.container.SplitterDistance + "\\";

            for (int c = 0; c < this.list_view.Columns.Count; c++)
            {
                Properties.Settings.Default.AllSettings
                    += this.list_view.Columns[c].Width + "\\";
            }
            for (int c = 0; c < this.grid_view.Columns.Count; c++)
            {
                Properties.Settings.Default.AllSettings
                    += this.grid_view.Columns[c].Width + "\\";
            }

            Properties.Settings.Default.Save();
        }
        private void PropertiesLoad()
        {
            if (Properties.Settings.Default.AllSettings != "")
            {
                int i = 0;
                string[] splits = Properties.Settings.Default.AllSettings.Split('\\');

                this.tool_container.SplitterDistance = int.Parse(splits[i++]);
                this.container.SplitterDistance = int.Parse(splits[i++]);

                for (int c = 0; c < this.list_view.Columns.Count; c++)
                {
                    this.list_view.Columns[c].Width = int.Parse(splits[i++]);
                }
                for (int c = 0; c < this.grid_view.Columns.Count; c++)
                {
                    this.grid_view.Columns[c].Width = int.Parse(splits[i++]);
                }
            }
        }
    }
}