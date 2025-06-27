using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using Autodesk.Revit.UI;
using Autodesk.Revit.DB;
using System.IO;
using System.Runtime.InteropServices;
using System.Security.Principal;
using Autodesk.Revit.DB.Visual;

namespace TRUEDREAMS
{
    public partial class Form3 : System.Windows.Forms.Form
    {
        private UIApplication uiapp;
        private UIDocument uidoc;
        private Document doc;
        private ExternalCommandData commandData;
        int chosen;
        string ppath;
        string newppath;
        string picname;

        // WindowsのAPI関数を使用してユーザーのログオンを行う
        [DllImport("advapi32.dll", SetLastError = true)]
        public static extern bool LogonUser(
           string lpszUsername,
           string lpszDomain,
           string lpszPassword,
           int dwLogonType,
           int dwLogonProvider,
           ref IntPtr phToken);

        // ハンドルを閉じる（ログアウト相当）
        [DllImport("kernel32.dll")]
        public extern static bool CloseHandle(IntPtr hToken);

        public Form3(ExternalCommandData commandData1)
        {
            InitializeComponent();

            // リモートフォルダへログイン
            string UserName = "";
            string MachineName = "";
            string Pw = "";
            const int LOGON32_PROVIDER_DEFAULT = 0;
            const int LOGON32_LOGON_NEW_CREDENTIALS = 9;
            IntPtr tokenHandle = new IntPtr(0);
            tokenHandle = IntPtr.Zero;
            bool returnValue = LogonUser(UserName, "", Pw, LOGON32_LOGON_NEW_CREDENTIALS, LOGON32_PROVIDER_DEFAULT, ref tokenHandle);

            WindowsIdentity w = new WindowsIdentity(tokenHandle);
            w.Impersonate();
            if (false == returnValue)
            {
                return;
            }

            string IPath =;

            commandData = commandData1;
            uiapp = commandData.Application;
            uidoc = uiapp.ActiveUIDocument;
            doc = uidoc.Document;

            // Revit のマテリアル情報をツリービューに追加
            Autodesk.Revit.DB.FilteredElementCollector filter = new FilteredElementCollector(doc).OfClass(typeof(Autodesk.Revit.DB.Material));

            foreach (Element materialElement in filter.ToElements())
            {
                Autodesk.Revit.DB.Material m = materialElement as Autodesk.Revit.DB.Material;
                TreeNode newNode = new TreeNode(m.Name);
                treeView2.Nodes.Add(newNode);
            }

            // 材質フォルダのツリービュー構築
            int i = 0;
            DirectoryInfo dir = new DirectoryInfo(IPath);
            foreach (DirectoryInfo dChild in dir.GetDirectories("*"))
            {
                TreeNode newNode = new TreeNode(dChild.Name);
                treeView1.Nodes.Add(newNode);
                DirectoryInfo dir2 = new DirectoryInfo(IPath + dChild.Name);
                int j = 0;
                foreach (DirectoryInfo dChild_d in dir2.GetDirectories("*"))
                {
                    TreeNode newNode2 = new TreeNode(dChild_d.Name);
                    treeView1.Nodes[i].Nodes.Add(newNode2);
                    DirectoryInfo dir3 = new DirectoryInfo(dChild_d.FullName);
                    foreach (DirectoryInfo dChild_d2 in dir3.GetDirectories("*"))
                    {
                        TreeNode newNode3 = new TreeNode(dChild_d2.Name);
                        treeView1.Nodes[i].Nodes[j].Nodes.Add(newNode3);
                    }
                    j++;
                }
                i++;    
            }
        }

        private void Form3_Load(object sender, EventArgs e)
        {
        }

        // 指定パスからサムネイル画像を取得
        private Bitmap GetThumbNail(string path)
        {
            Bitmap myBitmap = new Bitmap(path);
            Bitmap newBmp = new Bitmap(150, 150);
            Graphics g = Graphics.FromImage(newBmp);
            Pen p = new Pen(System.Drawing.Color.Black);

            int newWidth = 0;
            int newHeight = 0;
            int startX = 0;
            int startY = 0;

            if (myBitmap.Width > myBitmap.Height * 2)
            {
                newWidth = 150;
                newHeight = myBitmap.Height * 150 / myBitmap.Width;
                startX = (100 - newWidth) / 2;
                startY = (100 - newHeight) / 2;
            }
            else if (myBitmap.Height > myBitmap.Width * 2)
            {
                newWidth = myBitmap.Width * 150 / myBitmap.Height;
                newHeight = 150;
                startX = (100 - newWidth) / 2;
                startY = (100 - newHeight) / 2;
            }
            else
            {
                newWidth = 150;
                newHeight = 150;
                startX = (100 - newWidth) / 2;
                startY = (100 - newHeight) / 2;
            }

            // サムネイル画像を描画
            g.DrawImage(myBitmap, startX, startY, newWidth, newHeight);
            g.Dispose();
            myBitmap.Dispose();

            return newBmp;
        }
    

		private void TreeView1_AfterSelect(object sender, TreeViewEventArgs e)
		{
			// ユーザーがツリービューで項目を選択したときの処理

			/* 遠隔フォルダへのログイン設定 */
			string UserName = ""; // ユーザー名
			string MachineName = ""; // マシン名（IPまたはホスト名）
			string Pw = ""; // パスワード
			const int LOGON32_PROVIDER_DEFAULT = 0;
			const int LOGON32_LOGON_NEW_CREDENTIALS = 9;
			IntPtr tokenHandle = new IntPtr(0);
			tokenHandle = IntPtr.Zero;

			// Tokenの取得（ログイン）
			bool returnValue = LogonUser(UserName, "", Pw, LOGON32_LOGON_NEW_CREDENTIALS, LOGON32_PROVIDER_DEFAULT, ref tokenHandle);

			// 模擬ユーザーとして実行
			WindowsIdentity w = new WindowsIdentity(tokenHandle);
			w.Impersonate();
			if (!returnValue)
			{
				// ログイン失敗時の処理
				return;
			}

			string IPath = ; // 選択されたパスを指定する必要あり
			/* 遠隔フォルダへのログイン設定ここまで */

			// UI初期化：不要なラベルやテキストボックスを非表示
			label2.Visible = false;
			label8.Visible = false;
			label9.Visible = false;
			label4.Visible = false;
			label5.Visible = false;
			label10.Visible = false;
			textBox1.Visible = false;
			textBox2.Visible = false;
			textBox2.Text = "30.48";
			textBox1.Text = "30.48";
			groupBox1.Visible = false;
			label16.Visible = false;
			button2.Visible = false;
			textBox3.Visible = false;
			label6.Visible = false;

			try
			{
				if (treeView1.SelectedNode.Level == 2)
				{
					// フォルダ構造：カテゴリ/サブカテゴリ/マテリアル名の3階層
					DirectoryInfo dir2 = new DirectoryInfo(IPath + treeView1.SelectedNode.Parent.Parent.Text + @"\" + treeView1.SelectedNode.Parent.Text + @"\" + treeView1.SelectedNode.Text);

					FileInfo[] afi = dir2.GetFiles("*.*");
					string fileName;
					IList<FileInfo> list = new List<FileInfo>();

					// 対象ファイル（画像）の絞り込み
					for (int j = 0; j < afi.Length; j++)
					{
						fileName = afi[j].Name.ToLower();
						if (fileName.EndsWith(".jpg") || fileName.EndsWith(".png") || fileName.EndsWith(".jpeg") || fileName.EndsWith(".tif") || fileName.EndsWith(".tiff") || fileName.EndsWith(".bmp") || fileName.EndsWith(".jfif"))
						{
							list.Add(afi[j]);
						}
					}

					tabControl1.TabPages.Clear();
					int amount = list.Count;

					// UI要素配列の初期化
					PictureBox[] pic = new PictureBox[amount];
					TabPage[] tab = new TabPage[amount];
					Button[] btn = new Button[amount];
					Label[] lab = new Label[amount];
					Label[] scle = new Label[amount];
					Bitmap[] thumbnail = new Bitmap[amount];
					Image[] images = new Image[amount];

					int i = 0;
					int y = 20;

					tab[0] = new TabPage();
					foreach (FileInfo dChild_d in list)
					{
						// サムネイル画像表示エリア
						pic[i] = new PictureBox
						{
							Height = 150,
							Width = 150,
							Location = new System.Drawing.Point(20, y)
						};

						string[] subs = dChild_d.Name.Split('_');

						// ファイル名情報から取得したマテリアル情報の表示
						lab[i] = new Label
						{
							Text = subs[2] + "_" + subs[3] + "_" + subs[4] + "_" + subs[5],
							Font = new Font("微軟正黑體", 8),
							Height = 15,
							Width = 200,
							Location = new System.Drawing.Point(180, y + 55),
							AutoSize = false,
							Name = "lab" + i.ToString()
						};
						tab[0].Controls.Add(lab[i]);

						// スケール情報の表示
						scle[i] = new Label
						{
							Text = subs[1],
							Font = new Font("微軟正黑體", 8),
							Height = 15,
							Width = 60,
							Location = new System.Drawing.Point(425, y + 55),
							AutoSize = false,
							Name = "lab" + i.ToString()
						};
						tab[0].Controls.Add(scle[i]);

						// サムネイルの取得と表示
						thumbnail[i] = GetThumbNail(dir2.FullName + @"\" + dChild_d.Name);
						images[i] = thumbnail[i];
						pic[i].Image = images[i];
						pic[i].SizeMode = PictureBoxSizeMode.StretchImage;
						tab[0].Controls.Add(pic[i]);
						tab[0].UseVisualStyleBackColor = false;
						tab[0].BackColor = System.Drawing.Color.White;
						tab[0].AutoScroll = true;

						// 選択ボタンの作成
						btn[i] = new Button();
						btn[i].Text = "選此貼圖"; // 「このテクスチャを選択」
						btn[i].Font = new Font("微軟正黑體", 10);
						btn[i].Height = 30;
						btn[i].Width = 85;
						btn[i].Name = "btn" + i.ToString();
						btn[i].Location = new System.Drawing.Point(530, y + 50);
						btn[i].Click += new EventHandler(btn_Click);
						tab[0].Controls.Add(btn[i]);

						i++;
						y += 150;
					}

					// タブの設定とUI再表示
					tab[0].Text = treeView1.SelectedNode.Text;
					tabControl1.TabPages.Add(tab[0]);
					label16.Visible = true;
					button2.Visible = true;
					textBox3.Visible = true;
				}
			}
			catch
			{
				// エラー時メッセージ表示
				label6.Visible = true;
			}
		}
       

		   /// <summary>
			/// 選択された画像ボタンがクリックされた時の処理。
			/// 遠隔フォルダから選択された画像の情報を取得し、UI に反映する。
			/// </summary>
			void btn_Click(Object sender, EventArgs e)
			{
				/* リモートフォルダにログイン */
				string UserName = "";
				string MachineName = "";
				string Pw = "";
				const int LOGON32_PROVIDER_DEFAULT = 0;
				const int LOGON32_LOGON_NEW_CREDENTIALS = 9;
				IntPtr tokenHandle = new IntPtr(0);
				tokenHandle = IntPtr.Zero;

				// トークンにログイン情報を設定
				bool returnValue = LogonUser(UserName, "", Pw, LOGON32_LOGON_NEW_CREDENTIALS, LOGON32_PROVIDER_DEFAULT, ref tokenHandle);

				// ログインユーザーとして動作を模擬
				WindowsIdentity w = new WindowsIdentity(tokenHandle);
				w.Impersonate();
				if (false == returnValue)
				{
					// ログイン失敗時の処理
					return;
				}

				string IPath = ;
				/* リモートフォルダにログイン */

				// 選択されたノードのフォルダを取得
				DirectoryInfo dir2 = new DirectoryInfo(IPath + treeView1.SelectedNode.Parent.Parent.Text + @"\" + treeView1.SelectedNode.Parent.Text + @"\" + treeView1.SelectedNode.Text);

				FileInfo[] afi = dir2.GetFiles("*.*");

				string fileName;
				IList<FileInfo> list = new List<FileInfo>();

				// 画像ファイルをフィルタリング
				for (int j = 0; j < afi.Length; j++)
				{
					fileName = afi[j].Name.ToLower();
					if (fileName.EndsWith(".jpg") || fileName.EndsWith(".png") || fileName.EndsWith(".jpeg") || fileName.EndsWith(".tif") || fileName.EndsWith(".tiff") || fileName.EndsWith(".bmp") || fileName.EndsWith(".jfif"))
					{
						list.Add(afi[j]);
					}
				}

				Button temp = (Button)sender;
				string fid = temp.Name.Trim('b', 't', 'n');
				chosen = Int32.Parse(fid);
				ppath = list[chosen].FullName;

				string[] subs = list[chosen].Name.Split('_');

				// ラベルにファイル名の一部を表示
				label2.Text = subs[3] + "_" + subs[4] + "_" + subs[5];
				picname = list[chosen].Name;

				// 寸法の設定
				if (subs[1] == "無尺寸")
				{
					textBox2.Text = "30.48";
					textBox1.Text = "30.48";
				}
				else
				{
					string[] scale = subs[1].Split('X');
					textBox2.Text = scale[0];
					textBox1.Text = scale[1];
				}

				label2.Width = 200;
				label2.AutoSize = false;
				label2.AutoEllipsis = true;
				label2.Visible = true;
				label8.Visible = true;

				// UI 要素の表示切り替え
				if (label10.Visible == true)
				{
					label4.Visible = true;
					label5.Visible = true;
					textBox1.Visible = true;
					textBox2.Visible = true;
					groupBox1.Visible = true;
				}
			}

					/// <summary>
			/// マテリアルのレンダリングテクスチャパスを変更する処理
			/// </summary>
			void ChangeRenderingTexturePath(Document doc)
			{
				try
				{
					string mname = label10.Text; // 選択されたマテリアル名
					string texturePath = newppath; // 新しいテクスチャパス
					double X = double.Parse(textBox2.Text); // 幅（cm）
					double Y = double.Parse(textBox1.Text); // 高さ（cm）

					Autodesk.Revit.DB.FilteredElementCollector filter = new FilteredElementCollector(doc).OfClass(typeof(Autodesk.Revit.DB.Material));

					foreach (Element materialElement in filter.ToElements())
					{
						Autodesk.Revit.DB.Material m = materialElement as Autodesk.Revit.DB.Material;
						if (m.Name == mname)
						{
							using (Transaction t = new Transaction(doc, "Changing material texture path"))
							{
								t.Start();

								using (AppearanceAssetEditScope editScope = new AppearanceAssetEditScope(doc))
								{
									Asset editableAsset = editScope.Start(m.AppearanceAssetId);
									AssetProperty assetProperty = editableAsset.FindByName("generic_diffuse");

									Asset connectedAsset = assetProperty.GetSingleConnectedAsset() as Asset;

									if (connectedAsset == null)
									{
										assetProperty.AddConnectedAsset("UnifiedBitmapSchema");
										connectedAsset = assetProperty.GetSingleConnectedAsset();
									}

									if (connectedAsset.Name == "UnifiedBitmapSchema")
									{
										AssetPropertyString path_T = connectedAsset.FindByName(UnifiedBitmap.UnifiedbitmapBitmap) as AssetPropertyString;
										AssetPropertyDistance path_X = connectedAsset.FindByName(UnifiedBitmap.TextureRealWorldScaleX) as AssetPropertyDistance;
										AssetPropertyDistance path_Y = connectedAsset.FindByName(UnifiedBitmap.TextureRealWorldScaleY) as AssetPropertyDistance;
										AssetPropertyBoolean path_LOCK = connectedAsset.FindByName(UnifiedBitmap.TextureScaleLock) as AssetPropertyBoolean;

										path_LOCK.Value = false; // スケールロック解除

										if (path_T.IsValidValue(texturePath))
											path_T.Value = texturePath;

										if (path_X.IsValidValue(X / 2.54))
											path_X.Value = X / 2.54;

										if (path_Y.IsValidValue(Y / 2.54))
											path_Y.Value = Y / 2.54;
									}

									editScope.Commit(true);
								}
								TaskDialog.Show("套用材質貼圖成功", "專案材料 " + label10.Text + " \n材質貼圖已成功替換為 " + label2.Text);
								t.Commit();
								t.Dispose();
							}
							break;
						}
					}
				}
				catch
				{
					MessageBox.Show("抱歉，該材料無法被程式替換貼圖，請手動變更。");
				}
			}

			/// <summary>
			/// 「選擇貼圖」ボタンを押した際の処理
			/// </summary>
			private void Button1_Click(object sender, EventArgs e)
			{
				if (label2.Visible == false)
				{
					MessageBox.Show("未選擇材質貼圖！");
				}
				else if (label10.Visible == false)
				{
					MessageBox.Show("請選擇要賦予材質貼圖的專案材料！");
				}
				else if (textBox1.Text == "" || textBox2.Text == "")
				{
					MessageBox.Show("未輸入貼圖尺寸！");
				}
				else if (textBox3.Text == "")
				{
					MessageBox.Show("請選擇材質貼圖儲存路徑！(網路路徑會導致貼圖可能時常遺失！)");
				}
				else
				{
					// リモートフォルダへのログイン
					string UserName = "";
					string MachineName = "";
					string Pw = "";
					const int LOGON32_PROVIDER_DEFAULT = 0;
					const int LOGON32_LOGON_NEW_CREDENTIALS = 9;
					IntPtr tokenHandle = new IntPtr(0);
					tokenHandle = IntPtr.Zero;
					bool returnValue = LogonUser(UserName, "", Pw, LOGON32_LOGON_NEW_CREDENTIALS, LOGON32_PROVIDER_DEFAULT, ref tokenHandle);

					WindowsIdentity w = new WindowsIdentity(tokenHandle);
					w.Impersonate();
					if (false == returnValue)
					{
						return;
					}

					string IPath = ;

					// 新しいパスへファイルをコピー
					newppath = Path.Combine(textBox3.Text, picname);
					File.Copy(ppath, newppath, true);

					// テクスチャを適用
					ChangeRenderingTexturePath(doc);
				}
			}

			private void TabPage1_Click(object sender, EventArgs e)
			{
			}

			/// <summary>
			/// マテリアル選択時の処理
			/// </summary>
			private void TreeView2_AfterSelect(object sender, TreeViewEventArgs e)
			{
				if (treeView2.SelectedNode == null) { }
				else
				{
					label9.Visible = true;
					label10.Visible = true;
					label10.AutoSize = false;
					label10.AutoEllipsis = true;
					label10.Width = 170;
					label10.Text = treeView2.SelectedNode.Text;
				}

				if (label2.Visible == true)
				{
					label4.Visible = true;
					label5.Visible = true;
					textBox1.Visible = true;
					textBox2.Visible = true;
					groupBox1.Visible = true;
				}
			}

			private void label7_Click(object sender, EventArgs e)
			{
			}

			/// <summary>
			/// 保存先フォルダ選択ダイアログ表示
			/// </summary>
			private void button2_Click(object sender, EventArgs e)
			{
				FolderBrowserDialog path = new FolderBrowserDialog();
				path.ShowDialog();
				textBox3.Text = path.SelectedPath;
			}
