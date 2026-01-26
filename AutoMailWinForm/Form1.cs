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
			// [SỬA 1] Lấy số luồng từ ô nhập liệu thay vì số cố định
			int soLuongThread = (int)nudThreads.Value;

			// 2. CHUẨN BỊ GIAO DIỆN: Tạo sẵn các dòng "Đang chờ..."
			dtResults.Clear();
			List<MailTask> taskList = new List<MailTask>();

			for (int i = 0; i < rawList.Count; i++)
			{
				int stt = i + 1;
				// Thêm sẵn vào bảng để bạn thấy danh sách 20 mail
				dtResults.Rows.Add(stt, rawList[i], "", "Đang chờ lượt...");

				// Tạo task để chạy
				taskList.Add(new MailTask { Stt = stt, Email = rawList[i] });
			}

			btnStart.Enabled = false;
			if (btnExport != null) btnExport.Enabled = false; // Khóa nút xuất file
			btnStart.Text = $"Đang chạy {taskList.Count} mail...";

			await Task.Run(() =>
			{
				// --- CẤU HÌNH SỐ LUỒNG ---
				//Dùng biến soLuongThread vào cấu hình Parallel
				var options = new ParallelOptions { MaxDegreeOfParallelism = soLuongThread };

				Parallel.ForEach(taskList, options, (task) =>
				{
					// Truyền STT vào để code biết đang xử lý dòng nào
					XuLyMotTaiKhoan(task.Email, task.Stt);
				});
			});

			btnStart.Enabled = true;
			if (btnExport != null) btnExport.Enabled = true;
			btnStart.Text = "Bắt đầu chạy";
			MessageBox.Show("Hoàn thành!");
		}

		// Thêm tham số 'stt' vào hàm xử lý
		void XuLyMotTaiKhoan(string fullEmail, int stt)
		{
			HashSet<string> collectedLinks = new HashSet<string>();

			UpdateStatus(stt, "Đang khởi tạo...");

			var chromeOptions = new ChromeOptions();
			chromeOptions.AddArgument("--window-size=1000,800");

			// --- CẤU HÌNH TĂNG TỐC ĐỘ (CODE MỚI) ---
			chromeOptions.AddArgument("--headless=new"); // 1. Chạy ẩn

			// 2. Chặn tải hình ảnh (Giúp web load nhanh gấp đôi)
			chromeOptions.AddUserProfilePreference("profile.managed_default_content_settings.images", 2);
			chromeOptions.AddArgument("--blink-settings=imagesEnabled=false");

			// 3. Chế độ Eager: Chỉ cần hiện chữ là chạy, không chờ quảng cáo load xong
			chromeOptions.PageLoadStrategy = PageLoadStrategy.Eager;

			// 4. Tắt các tính năng thừa thãi
			chromeOptions.AddArgument("--disable-extensions");
			chromeOptions.AddArgument("--disable-gpu");
			chromeOptions.AddArgument("--disable-popup-blocking");

			using (var driver = new ChromeDriver(chromeOptions))
			{
				try
				{
					driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(5);
					WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(40)); // Tăng thời gian chờ tổng

					// --- LOGIN ---
					driver.Navigate().GoToUrl("https://hcmail.xyz");
					wait.Until(ExpectedConditions.ElementIsVisible(By.CssSelector("input[type='password']"))).SendKeys("1" + OpenQA.Selenium.Keys.Enter);

					var btnMoi = wait.Until(ExpectedConditions.ElementToBeClickable(By.XPath("//i[contains(@class, 'fa-plus')]/parent::div")));
					((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].click();", btnMoi);

					// 1. Tách user và domain
					string username = fullEmail.Split('@')[0];
					string targetDomain = fullEmail.Split('@')[1];

					var inputUser = wait.Until(ExpectedConditions.ElementIsVisible(By.Name("user")));
					inputUser.Clear();
					// 2. Chỉ điền Username
					inputUser.SendKeys(username);

					// 3. Chọn đúng Domain từ menu
					try
					{
						var domainSelect = driver.FindElement(By.Name("domain"));
						((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].click();", domainSelect);
						Thread.Sleep(200);

						// Click vào dòng chứa đúng tên miền (hcmail.xyz hoặc hvcmail.online...)
						var domainOption = driver.FindElement(By.XPath($"//a[contains(text(), '{targetDomain}')]"));
						((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].click();", domainOption);
					}
					catch { }



					var btnTao = driver.FindElement(By.XPath("//input[@type='submit']"));
					((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].click();", btnTao);

					// --- ĐỌC 5 THƯ ---
					UpdateStatus(stt, "Đang chờ thư về...");
					// Chờ danh sách hiện đủ 5 cái
					wait.Until(d => d.FindElements(By.XPath("//div[@data-id]")).Count >= 5);

					// Vòng lặp từ 4 về 0 (để lấy đúng thứ tự Mail 1 -> Mail 5)
					for (int i = 0; i < 5; i++)
					{
						int mailLabel = i + 1;
						bool isSuccess = false;

						// === CƠ CHẾ THỬ LẠI (RETRY) ===
						// Nếu lỗi ở thư này, thử lại tối đa 3 lần
						for (int retry = 0; retry < 3; retry++)
						{
							try
							{
								UpdateStatus(stt, $"Đang đọc thư {mailLabel}/5 (Lần {retry + 1})...");

								// Lấy lại danh sách mail (tránh lỗi Stale Element)
								wait.Until(d => d.FindElements(By.XPath("//div[@data-id]")).Count >= 5);
								var currentMails = driver.FindElements(By.XPath("//div[@data-id]"));

								// Scroll xuống để hiện mail
								// Gộp lệnh Scroll và Click làm 1, bấm cực nhanh
								((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].scrollIntoView(true); arguments[0].click();", currentMails[i]);

								// TÌM LINK
								string finalLink = "";
								bool isDuplicate = true; // Giả định là trùng, cần tìm cái mới

								// Quét iframe tìm link trong 10 giây
								for (int attempt = 0; attempt < 20; attempt++)
								{
									try
									{
										var iframes = driver.FindElements(By.TagName("iframe"));
										foreach (var iframe in iframes.Reverse())
										{
											string raw = iframe.GetAttribute("srcdoc");
											if (!string.IsNullOrEmpty(raw))
											{
												string decoded = System.Net.WebUtility.HtmlDecode(raw);
												Match m = Regex.Match(decoded, @"href\s*=\s*[""'](https?://[^""']+)[""']");
												string tempLink = m.Success ? m.Groups[1].Value : "";

												if (string.IsNullOrEmpty(tempLink))
												{
													Match m2 = Regex.Match(decoded, @"https?://[^\s""'<]+");
													if (m2.Success) tempLink = m2.Value;
												}

												if (!string.IsNullOrEmpty(tempLink))
												{
													// Kiểm tra trùng lặp
													if (!collectedLinks.Contains(tempLink))
													{
														finalLink = tempLink;
														isDuplicate = false;
														// Xóa iframe để dọn dẹp
														((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].remove();", iframe);
														break;
													}
												}
											}
										}
									}
									catch { }

									if (!isDuplicate) break; // Tìm thấy link mới -> Thoát vòng lặp chờ
									Thread.Sleep(300);
								}

								if (!string.IsNullOrEmpty(finalLink))
								{
									collectedLinks.Add(finalLink);
									UpdateResult(stt, finalLink, mailLabel);
								}
								else
								{
									UpdateResult(stt, "Không tìm thấy link", mailLabel);
								}

								// Thành công -> Đánh dấu thoát vòng lặp Retry
								isSuccess = true;

								// Quay về Inbox chuẩn bị cho thư sau
								driver.Navigate().GoToUrl("https://hcmail.xyz/mailbox");
								break;
							}
							catch
							{
								// Nếu lỗi, quay về Inbox và thử lại (vòng lặp retry sẽ chạy tiếp)
								driver.Navigate().GoToUrl("https://hcmail.xyz/mailbox");
								Thread.Sleep(2000);
							}
						}

						// Nếu thử 3 lần mà vẫn thất bại hoàn toàn
						if (!isSuccess)
						{
							UpdateResult(stt, "Lỗi đọc thư (Skip)", mailLabel);
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
							string linkCu = row["Link Lấy Được"].ToString();
							string linkMoi = $"[Mail {mailIndex}]: {link}";

							if (string.IsNullOrEmpty(linkCu) || mailIndex == 1)
								row["Link Lấy Được"] = linkMoi;
							else
								row["Link Lấy Được"] = linkCu + Environment.NewLine + linkMoi;

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
					lines.Add("STT,Email,Link 1,Link 2,Link 3,Link 4,Link 5,Trạng thái");

					foreach (DataRow row in dtResults.Rows)
					{
						// Lấy chuỗi gốc trong ô (đang chứa cả 5 link và xuống dòng)
						string rawLinks = row["Link Lấy Được"].ToString();

						// Tách từng dòng ra
						string[] linkArray = rawLinks.Split(new[] { Environment.NewLine }, StringSplitOptions.RemoveEmptyEntries);

						// Khai báo 5 biến chứa link, mặc định là rỗng
						string l1 = "", l2 = "", l3 = "", l4 = "", l5 = "";

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
						}

						// 2. GHI DÒNG DỮ LIỆU VỚI 5 CỘT
						// Cấu trúc: STT, Email, L1, L2, L3, L4, L5, Trạng Thái
						string line = $"{row["STT"]},{row["Email"]},{l1},{l2},{l3},{l4},{l5},{row["Trạng Thái"]}";
						lines.Add(line);
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