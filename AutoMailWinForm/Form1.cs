using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using SeleniumExtras.WaitHelpers;
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using WebDriverManager;
using WebDriverManager.DriverConfigs.Impl;
using WebDriverManager.Helpers;
namespace AutoMailWinForm
{
	public partial class Form1 : Form
	{
		DataTable dtResults;
		NumericUpDown nudThreads;
		Label lblThreads;
		public Form1()
		{
			InitializeComponent();
			SetupGrid();
			SetupCustomControls();
		}
		void SetupCustomControls()
		{
			// Tạo nhãn "Số luồng:"
			lblThreads = new Label();
			lblThreads.Text = "Số luồng:";
			lblThreads.AutoSize = true;
			lblThreads.Location = new Point(btnStart.Right + 20, btnStart.Top + 5);
			this.Controls.Add(lblThreads);

			// Tạo ô nhập số
			nudThreads = new NumericUpDown();
			nudThreads.Minimum = 1;
			nudThreads.Maximum = 50;  // Max 50 luồng
			nudThreads.Value = 5;     // Mặc định 5
			nudThreads.Width = 50;
			nudThreads.Location = new Point(lblThreads.Right + 5, btnStart.Top + 2);
			this.Controls.Add(nudThreads);
		}
		void SetupGrid()
		{
			dtResults = new DataTable();
			dtResults.Columns.Add("STT");
			dtResults.Columns.Add("Email");
			dtResults.Columns.Add("Link Lấy Được");
			dtResults.Columns.Add("Trạng Thái");

			gridResult.DataSource = dtResults;

			// 1. CHÌA KHÓA: Cho phép các cột tự động giãn ra để lấp đầy toàn bộ chiều rộng bảng
			gridResult.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;

			// 2. Cấu hình tỉ lệ độ rộng giữa các cột (Tổng FillWeight là 100)
			// Cột STT chiếm 10% chiều rộng
			gridResult.Columns[0].FillWeight = 5;
			// Cột Email chiếm 25%
			gridResult.Columns[1].FillWeight = 20;
			// Cột Link chiếm 50% (ưu tiên rộng nhất vì dữ liệu dài)
			gridResult.Columns[2].FillWeight = 55;
			// Cột Trạng Thái chiếm 15%
			gridResult.Columns[3].FillWeight = 20;

			// 3. Giữ cấu hình tự động xuống dòng và co giãn chiều cao dòng
			gridResult.Columns[2].DefaultCellStyle.WrapMode = DataGridViewTriState.True;
			gridResult.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;

			// (Tùy chọn) Ngăn người dùng kéo chỉnh kích thước cột thủ công làm hỏng giao diện
			gridResult.AllowUserToResizeColumns = true;
		}

		private async void btnStart_Click(object sender, EventArgs e)
		{
			// Dòng này sẽ kiểm tra Chrome trên máy, tải driver khớp version về.
			await Task.Run(() => new DriverManager().SetUpDriver(
			new ChromeConfig(),
			VersionResolveStrategy.MatchingBrowser
));
			// 1. Lấy danh sách từ ô nhập liệu
			List<string> rawList = txtInput.Lines
								   .Where(x => !string.IsNullOrWhiteSpace(x))
								   .Select(x => x.Trim())
								   .ToList();

			if (rawList.Count == 0)
			{
				MessageBox.Show("Bạn chưa nhập email nào cả!");
				return;
			}

			// Lấy số luồng từ ô nhập
			int soLuongThread = (int)nudThreads.Value;

			// 2. CHUẨN BỊ GIAO DIỆN
			dtResults.Clear();
			List<MailTask> taskList = new List<MailTask>();

			for (int i = 0; i < rawList.Count; i++)
			{
				int stt = i + 1;
				dtResults.Rows.Add(stt, rawList[i], "", "Đang chờ lượt...");
				taskList.Add(new MailTask { Stt = stt, Email = rawList[i] });
			}

			// Khóa các nút
			btnStart.Enabled = false;
			if (btnExport != null) btnExport.Enabled = false;
			nudThreads.Enabled = false;

			btnStart.Text = $"Đang chạy {taskList.Count} mail...";

			// --- [BẮT ĐẦU CƠ CHẾ SEMAPHORE (GHẾ NGỒI)] ---

			// Tạo 'soLuongThread' cái ghế. Ví dụ: 5 ghế.
			using (SemaphoreSlim semaphore = new SemaphoreSlim(soLuongThread))
			{
				// Tạo danh sách các tác vụ (Task) cho TẤT CẢ email
				var tasks = taskList.Select(async task =>
				{
					// 1. XIN GHẾ: Chờ đến khi có ghế trống
					await semaphore.WaitAsync();
					try
					{
						// 2. CÓ GHẾ RỒI -> CHẠY NGAY
						// Chạy hàm xử lý trong một luồng riêng biệt để không đơ giao diện
						await Task.Run(() =>
						{
							XuLyMotTaiKhoan(task.Email, task.Stt);
						});
					}
					finally
					{
						// 3. TRẢ GHẾ: Dù chạy xong hay lỗi cũng phải trả ghế
						// Để người tiếp theo (đang đợi ở bước 1) được vào ngay lập tức
						semaphore.Release();
					}
				});

				// Đợi cho đến khi tất cả các email đều chạy xong
				await Task.WhenAll(tasks);
			}

			// --- [KẾT THÚC CHẠY] ---

			// Mở lại nút
			btnStart.Enabled = true;
			if (btnExport != null) btnExport.Enabled = true;
			nudThreads.Enabled = true;

			btnStart.Text = "Bắt đầu chạy";
			MessageBox.Show("Hoàn thành tất cả!");

			// Tự động xuất file (nếu muốn)
			// XuatFileCSV(); 
		}

        // Thêm tham số 'stt' vào hàm xử lý
        void XuLyMotTaiKhoan(string fullEmail, int stt)
        {
            // List lưu kết quả
            HashSet<string> collectedLinks = new HashSet<string>();
            UpdateStatus(stt, "Đang khởi tạo...");

            var chromeOptions = new ChromeOptions();
            chromeOptions.AddArgument("--window-size=1400,1000");
            chromeOptions.AddArgument("--headless=new");
            chromeOptions.AddUserProfilePreference("profile.managed_default_content_settings.images", 2);
            chromeOptions.AddArgument("--blink-settings=imagesEnabled=false");
            chromeOptions.PageLoadStrategy = PageLoadStrategy.Eager;
            chromeOptions.AddArgument("--disable-extensions");
            chromeOptions.AddArgument("--disable-gpu");
            chromeOptions.AddArgument("--disable-popup-blocking");

            using (var driver = new ChromeDriver(chromeOptions))
            {
                try
                {
                    driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(5);
                    WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(40));

                    // =========================================================
                    // 1. LOGIN & TẠO MAIL (GIỮ NGUYÊN)
                    // =========================================================
                    driver.Navigate().GoToUrl("https://tmailstore.com");

                    try
                    {
                        var passInput = wait.Until(ExpectedConditions.ElementIsVisible(By.Id("password")));
                        passInput.SendKeys("Trongtrieu970819@" + OpenQA.Selenium.Keys.Enter);
                        Thread.Sleep(2000);
                    }
                    catch { }

                    var btnMoi = wait.Until(ExpectedConditions.ElementToBeClickable(By.XPath("//div[contains(text(), 'Mới')]/parent::div")));
                    ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].click();", btnMoi);

                    string username = fullEmail.Split('@')[0];
                    var inputUser = wait.Until(ExpectedConditions.ElementIsVisible(By.Id("user")));
                    inputUser.Clear();
                    inputUser.SendKeys(username);

                    var btnTao = driver.FindElement(By.Id("create"));
                    ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].click();", btnTao);

                    // =========================================================
                    // 2. ĐỌC LINK (PHƯƠNG PHÁP KHÔNG CẦN CLICK)
                    // =========================================================
                    UpdateStatus(stt, "Đang chờ thư về...");

                    try
                    {
                        // Chờ danh sách bên trái hiện ra (để đảm bảo mail đã về)
                        wait.Until(d => d.FindElements(By.CssSelector(".messages > div[data-id]")).Count > 0);
                    }
                    catch
                    {
                        UpdateStatus(stt, "Timeout: Không có thư nào!");
                        return;
                    }

                    // [BƯỚC 1]: Lấy danh sách ID theo đúng thứ tự hiển thị (từ trên xuống dưới)
                    // Việc này giúp ta biết đâu là Mail 1, Mail 2...
                    var mailRows = driver.FindElements(By.CssSelector(".messages > div[data-id]"));
                    List<string> orderedIds = new List<string>();
                    foreach (var row in mailRows)
                    {
                        string id = row.GetAttribute("data-id");
                        if (!string.IsNullOrEmpty(id)) orderedIds.Add(id);
                    }

                    // Chỉ lấy 10 thư đầu
                    int count = Math.Min(10, orderedIds.Count);

                    // [BƯỚC 2]: Dùng vòng lặp quét qua từng ID để lấy nội dung
                    // Lưu ý: Nội dung nằm bên phải, trong các thẻ div có id="message-XXX"
                    for (int i = 0; i < count; i++)
                    {
                        int mailLabel = i + 1;
                        string targetId = orderedIds[i]; // Ví dụ: "36"
                        string targetDivId = "message-" + targetId; // Ví dụ: "message-36"

                        UpdateStatus(stt, $"Đang trích xuất Mail {mailLabel} (ID:{targetId})...");

                        try
                        {
                            // Tìm thẻ div chứa nội dung thư (dù nó đang ẩn display:none vẫn tìm được)
                            // Dựa vào HTML bạn gửi: <div x-show="id === 36" id="message-36" ...>
                            var contentDiv = driver.FindElement(By.Id(targetDivId));

                            // Tìm thẻ iframe nằm trong div đó
                            var iframe = contentDiv.FindElement(By.TagName("iframe"));

                            // [TUYỆT CHIÊU]: Lấy thẳng mã nguồn HTML trong thuộc tính 'srcdoc'
                            // Không cần SwitchTo().Frame() -> Tránh được mọi lỗi kẹt
                            string rawHtml = iframe.GetAttribute("srcdoc");

                            // Giải mã HTML (vì srcdoc bị mã hóa ký tự đặc biệt)
                            string decodedHtml = System.Net.WebUtility.HtmlDecode(rawHtml);

                            // Dùng Regex tìm link trong đoạn HTML vừa lấy
                            Match m = Regex.Match(decodedHtml, @"https://g4b\.giftee\.biz/giftee_boxes/[a-zA-Z0-9\-]+");

                            if (m.Success)
                            {
                                string link = m.Value;
                                collectedLinks.Add(link);
                                UpdateResult(stt, link, mailLabel);
                            }
                            else
                            {
                                UpdateResult(stt, "Không có link", mailLabel);
                            }
                        }
                        catch (Exception)
                        {
                            UpdateResult(stt, "Lỗi/Mail rỗng", mailLabel);
                        }
                    }

                    UpdateStatus(stt, "Hoàn tất!");
                }
                catch (Exception ex)
                {
                    UpdateStatus(stt, "Lỗi: " + ex.Message);
                }
            }
        }
        // --- CẬP NHẬT GIAO DIỆN THEO STT (ĐỊNH DANH DÒNG CHÍNH XÁC) ---
        void UpdateStatus(int stt, string status)
		{
			try
			{
				this.Invoke(new Action(() =>
				{
					foreach (DataRow row in dtResults.Rows)
					{
						// Tìm đúng dòng có STT khớp với luồng đang chạy
						if (Convert.ToInt32(row["STT"]) == stt)
						{
							row["Trạng Thái"] = status;
							break;
						}
					}
				}));
			}
			catch { }
		}

		void UpdateResult(int stt, string link, int mailIndex)
		{
			try
			{
				this.Invoke(new Action(() =>
				{
					foreach (DataRow row in dtResults.Rows)
					{
						if (Convert.ToInt32(row["STT"]) == stt)
						{
							// 1. Lấy nội dung hiện tại trong ô
							string currentText = row["Link Lấy Được"].ToString();

							// 2. Dùng SortedDictionary để lưu trữ
							// Key = Số thứ tự mail (1, 2, 3...)
							// Value = Link
							// Tác dụng: Tự động sắp xếp từ bé đến lớn và KHÔNG cho phép trùng Key
							SortedDictionary<int, string> linkMap = new SortedDictionary<int, string>();

							// 3. Phân tích (Parse) nội dung cũ để nhét lại vào Dictionary (để không bị mất dữ liệu cũ)
							if (!string.IsNullOrEmpty(currentText))
							{
								var lines = currentText.Split(new[] { Environment.NewLine }, StringSplitOptions.RemoveEmptyEntries);
								foreach (var line in lines)
								{
									// Dùng Regex để tách lấy số thứ tự mail: "[Mail 1]: https..."
									Match m = Regex.Match(line, @"\[Mail (\d+)\]: (.*)");
									if (m.Success)
									{
										int idx = int.Parse(m.Groups[1].Value);
										string content = m.Groups[2].Value;

										// Chỉ thêm nếu chưa có (đề phòng dữ liệu rác cũ)
										if (!linkMap.ContainsKey(idx))
										{
											linkMap.Add(idx, content);
										}
									}
								}
							}

							// 4. CẬP NHẬT LINK MỚI
							// Nếu mailIndex này đã tồn tại -> Ghi đè (Sửa lỗi lặp lại)
							// Nếu chưa có -> Thêm mới (Sửa lỗi lộn xộn vì Dictionary tự xếp)
							if (linkMap.ContainsKey(mailIndex))
							{
								linkMap[mailIndex] = link;
							}
							else
							{
								linkMap.Add(mailIndex, link);
							}

							// 5. Ghép lại thành chuỗi hiển thị chuẩn đẹp
							List<string> finalLines = new List<string>();
							foreach (var kvp in linkMap)
							{
								finalLines.Add($"[Mail {kvp.Key}]: {kvp.Value}");
							}

							row["Link Lấy Được"] = string.Join(Environment.NewLine, finalLines);

							// Auto resize để dòng giãn ra cho đẹp
							gridResult.AutoResizeRows(DataGridViewAutoSizeRowsMode.AllCells);
							break;
						}
					}
				}));
			}
			catch { }
		}

		void XuatFileCSV()
		{
			// Nếu đang ở luồng phụ thì gọi về luồng chính
			if (this.InvokeRequired) { this.Invoke(new Action(XuatFileCSV)); return; }

			try
			{
				SaveFileDialog sfd = new SaveFileDialog();
				sfd.Filter = "CSV File (*.csv)|*.csv";
				sfd.FileName = "KetQua_" + DateTime.Now.ToString("ddMMyyyy_HHmm") + ".csv";
				sfd.Title = "Lưu kết quả";

				if (sfd.ShowDialog() == DialogResult.OK)
				{
					List<string> lines = new List<string>();

                    // 1. TẠO HEADER VỚI 5 CỘT LINK RIÊNG BIỆT
                    lines.Add("STT,Email,Link 1,Link 2,Link 3,Link 4,Link 5,Link 6,Link 7,Link 8,Link 9,Link 10,Trạng thái");
                    foreach (DataRow row in dtResults.Rows)
					{
						// Lấy chuỗi gốc trong ô (đang chứa cả 5 link và xuống dòng)
						string rawLinks = row["Link Lấy Được"].ToString();

						// Tách từng dòng ra
						string[] linkArray = rawLinks.Split(new[] { Environment.NewLine }, StringSplitOptions.RemoveEmptyEntries);

                        // Khai báo 5 biến chứa link, mặc định là rỗng
                        string l1 = "", l2 = "", l3 = "", l4 = "", l5 = "", l6 = "", l7 = "", l8 = "", l9 = "", l10 = "";
                        // Duyệt qua từng dòng để nhét vào đúng biến
                        foreach (string item in linkArray)
						{
							// Item có dạng: "[Mail 1]: https://..." -> Cần cắt bỏ cái "[Mail 1]: " đi cho sạch
							string urlOnly = "";
							if (item.Contains("]: "))
								urlOnly = item.Split(new[] { "]: " }, StringSplitOptions.None)[1].Trim();
							else
								urlOnly = item; // Dự phòng

							if (item.Contains("[Mail 1]")) l1 = urlOnly;
							if (item.Contains("[Mail 2]")) l2 = urlOnly;
							if (item.Contains("[Mail 3]")) l3 = urlOnly;
							if (item.Contains("[Mail 4]")) l4 = urlOnly;
							if (item.Contains("[Mail 5]")) l5 = urlOnly;
                            if (item.Contains("[Mail 6]")) l6 = urlOnly;
                            if (item.Contains("[Mail 7]")) l7 = urlOnly;
                            if (item.Contains("[Mail 8]")) l8 = urlOnly;
                            if (item.Contains("[Mail 9]")) l9 = urlOnly;
                            if (item.Contains("[Mail 10]")) l10 = urlOnly;
                        }

                        // 2. GHI DÒNG DỮ LIỆU VỚI 5 CỘT
                        // Cấu trúc: STT, Email, L1, L2, L3, L4, L5, Trạng Thái
                        string line = $"{row["STT"]},{row["Email"]},{l1},{l2},{l3},{l4},{l5},{l6},{l7},{l8},{l9},{l10},{row["Trạng Thái"]}"; lines.Add(line);
					}

					// Ghi file UTF-8
					File.WriteAllLines(sfd.FileName, lines, System.Text.Encoding.UTF8);

					MessageBox.Show("Xuất file thành công! Hãy mở file bằng Excel.");
				}
			}
			catch (Exception ex) { MessageBox.Show("Lỗi xuất file: " + ex.Message); }
		}

		private void richTextBox1_TextChanged(object sender, EventArgs e) { }
		private void Form1_Load(object sender, EventArgs e) { }

		private void btnExport_Click(object sender, EventArgs e)
		{
			XuatFileCSV();
		}
	}

	// Class nhỏ để lưu thông tin công việc
	public class MailTask
	{
		public int Stt { get; set; }
		public string Email { get; set; }
	}
}