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
			await Task.Run(() => new DriverManager().SetUpDriver(new ChromeConfig()));
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
			// Giữ HashSet để lọc link trùng giữa các mail (Link rác, footer...)
			HashSet<string> collectedLinks = new HashSet<string>();

			UpdateStatus(stt, "Đang khởi tạo...");

			var chromeOptions = new ChromeOptions();
			chromeOptions.AddArgument("--window-size=1000,800");
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

					// --- LOGIN ---
					driver.Navigate().GoToUrl("https://hcmail.xyz");
					wait.Until(ExpectedConditions.ElementIsVisible(By.CssSelector("input[type='password']"))).SendKeys("1" + OpenQA.Selenium.Keys.Enter);

					var btnMoi = wait.Until(ExpectedConditions.ElementToBeClickable(By.XPath("//i[contains(@class, 'fa-plus')]/parent::div")));
					((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].click();", btnMoi);

					string username = fullEmail.Split('@')[0];
					string targetDomain = fullEmail.Split('@')[1];

					var inputUser = wait.Until(ExpectedConditions.ElementIsVisible(By.Name("user")));
					inputUser.Clear();
					inputUser.SendKeys(username);

					try
					{
						var domainSelect = driver.FindElement(By.Name("domain"));
						((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].click();", domainSelect);
						Thread.Sleep(200);
						var domainOption = driver.FindElement(By.XPath($"//a[contains(text(), '{targetDomain}')]"));
						((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].click();", domainOption);
					}
					catch { }

					var btnTao = driver.FindElement(By.XPath("//input[@type='submit']"));
					((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].click();", btnTao);

					// --- ĐỌC THƯ (LOGIC CHUẨN) ---
					UpdateStatus(stt, "Đang chờ thư về...");

					// [SỬA 1]: Chỉ cần chờ > 0 (Có thư là chạy, không ép chờ đủ 5 để tránh Timeout)
					try
					{
						wait.Until(d => d.FindElements(By.XPath("//div[@data-id]")).Count > 0);
					}
					catch
					{
						UpdateStatus(stt, "Không có thư nào (Timeout)!");
						return;
					}

					var currentMails = driver.FindElements(By.XPath("//div[@data-id]"));

					// [SỬA 2]: Tính toán số lượng cần đọc (Max 5 hoặc ít hơn)
					int countToRead = Math.Min(5, currentMails.Count);

					// Chạy vòng lặp từ 0 -> countToRead
					// i=0 tương ứng với phần tử đầu tiên trong HTML (Là Mới nhất theo phân tích logic trên)
					for (int i = 0; i < countToRead; i++)
					{
						int mailLabel = i + 1;
						bool isSuccess = false;

						for (int retry = 0; retry < 2; retry++)
						{
							try
							{
								UpdateStatus(stt, $"Đọc thư {mailLabel} (Lần {retry + 1})...");

								driver.Navigate().GoToUrl("https://hcmail.xyz/mailbox");
								wait.Until(d => d.FindElements(By.XPath("//div[@data-id]")).Count > 0);
								currentMails = driver.FindElements(By.XPath("//div[@data-id]"));

								// Kiểm tra an toàn index
								if (i >= currentMails.Count) break;

								// [CHỐT HẠ]: Click thẳng vào index i (0, 1, 2...)
								// Vì công thức đảo ngược đã sai, nên công thức xuôi này BẮT BUỘC ĐÚNG.
								((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].scrollIntoView(true); arguments[0].click();", currentMails[i]);

								// --- TÌM LINK ---
								string finalLink = "";

								// Quét iframe tìm link
								for (int attempt = 0; attempt < 20; attempt++)
								{
									try
									{
										var iframes = driver.FindElements(By.TagName("iframe"));
										foreach (var iframe in iframes.Reverse()) // Ưu tiên iframe cuối (nội dung chính)
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
													// Kiểm tra trùng: Nếu link này chưa từng có trong lịch sử của email này -> Lấy
													if (!collectedLinks.Contains(tempLink))
													{
														finalLink = tempLink;
														// Xóa iframe để nhẹ máy
														((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].remove();", iframe);
														break;
													}
												}
											}
										}
									}
									catch { }

									if (!string.IsNullOrEmpty(finalLink)) break;
									Thread.Sleep(250);
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

								isSuccess = true;
								break;
							}
							catch
							{
								driver.Navigate().GoToUrl("https://hcmail.xyz/mailbox");
								Thread.Sleep(1000);
							}
						}

						if (!isSuccess) UpdateResult(stt, "Lỗi đọc", mailLabel);
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