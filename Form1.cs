using System.Data;
using System.Diagnostics;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using MailKit;
using MailKit.Net.Imap;
using Microsoft.Playwright;

namespace peymyntReg;

public partial class RegForm : Form
{
    public RegForm()
    {
        InitializeComponent();
    }

    private const int WM_COPYDATA = 0x004A;

    private bool is911Applied;
    private string apply911String = "";

    public struct COPYDATASTRUCT
    {
        public IntPtr dwData;
        public int cData;

        [MarshalAs(UnmanagedType.LPStr)] public string lpData;
    }

    protected override void DefWndProc(ref Message m)
    {
        switch (m.Msg)
        {
            case WM_COPYDATA:
                var cds = new COPYDATASTRUCT();
                var t = cds.GetType();
                cds = (COPYDATASTRUCT)m.GetLParam(t);
                apply911String = cds.lpData;
                is911Applied = true;
                break;
            default:
                base.DefWndProc(ref m);
                break;
        }
    }

    private int Apply911Proxy()
    {
        var rd = new Random();
        var port = rd.Next(4000, 4100);
        var p = Process.Start(@"D:\911\ProxyTool\AutoProxyTool.exe",
            $"-changeproxy/US -proxyport={port} -hwnd={Handle}");
        return port;
    }

    private static DataTable CsvToDataTable(string filePath, int n)
    {
        var dt = new DataTable();
        var reader = new StreamReader(filePath, Encoding.Default, false);
        var m = 0;
        while (!reader.EndOfStream)
        {
            m = m + 1;
            var str = reader.ReadLine();
            var split = str.Split(',');
            if (m == n)
            {
                for (var c = 0; c < split.Length; c++)
                {
                    var column = new DataColumn();
                    column.DataType = Type.GetType("System.String");
                    column.ColumnName = split[c];
                    if (dt.Columns.Contains(split[c]))
                    {
                        column.ColumnName = split[c] + c;
                    }

                    dt.Columns.Add(column);
                }
            }

            if (m >= n + 1)
            {
                var dr = dt.NewRow();
                for (int i = 0; i < split.Length; i++)
                {
                    dr[i] = split[i];
                }

                dt.Rows.Add(dr);
            }
        }

        reader.Close();
        return dt;
    }

    private static void DataTaleToCsv(DataTable dt, string filePath)
    {
        if (dt == null || dt.Rows.Count == 0)
        {
            return;
        }

        var strBufferLine = "";
        var streamWriter = new StreamWriter(filePath, false, Encoding.Default);
        foreach (DataColumn col in dt.Columns)
        {
            strBufferLine += col.ColumnName + ",";
        }

        strBufferLine = strBufferLine.Substring(0, strBufferLine.Length - 1);
        streamWriter.WriteLine(strBufferLine);
        for (int i = 0; i < dt.Rows.Count; i++)
        {
            strBufferLine = "";
            for (int j = 0; j < dt.Columns.Count; j++)
            {
                if (j > 0)
                {
                    strBufferLine += ",";
                }

                strBufferLine += dt.Rows[i][j].ToString().Replace(",", "");
            }

            streamWriter.WriteLine(strBufferLine);
        }

        streamWriter.Close();
    }

    private static IMailFolder GetJunkFolder(ImapClient client)
    {
        var personal = client.GetFolder(client.PersonalNamespaces[0]);
        foreach (var folder in personal.GetSubfolders())
        {
            if (folder.Name == "Junk")
            {
                return folder;
            }
        }

        return null;
    }

    private string CheckEmailForURL(string[] emailInfo)
    {
        var isChecked = false;
        var startTime = new DateTimeOffset(DateTime.UtcNow).ToUnixTimeSeconds();
        var client = new ImapClient();
        client.Connect("outlook.office365.com", 993, true);
        client.Authenticate(emailInfo[0], emailInfo[1]);

        var inbox = client.Inbox;
        var junk = GetJunkFolder(client);

        var verifyUrl = "";
        var regex = new Regex(@"https://app.peymynt.com/verify\?\S*");
        while (!isChecked)
        {
            inbox.Open(FolderAccess.ReadOnly);
            for (var inboxIndex = 0; inboxIndex < inbox.Count; inboxIndex++)
            {
                var message = inbox.GetMessage(inboxIndex);
                if (message.From.Mailboxes.ToList()[0].Address == "admin@mailer.peymynt.com")
                {
                    if (message.Subject == "Confirm your Peymynt identity email address")
                    {
                        verifyUrl = regex.Match(message.HtmlBody).Value;
                        verifyUrl = verifyUrl.Substring(0, verifyUrl.Length - 1);
                        isChecked = true;
                        break;
                    }
                }
            }

            junk.Open(FolderAccess.ReadOnly);
            for (int junkIndex = 0; junkIndex < junk.Count; junkIndex++)
            {
                var message = junk.GetMessage(junkIndex);
                if (message.From.Mailboxes.ToList()[0].Address == "admin@mailer.peymynt.com")
                {
                    if (message.Subject == "Confirm your Peymynt identity email address")
                    {
                        verifyUrl = regex.Match(message.HtmlBody).Value;
                        verifyUrl = verifyUrl.Replace("&amp;", "&");
                        verifyUrl = verifyUrl.Substring(0, verifyUrl.Length - 1);
                        isChecked = true;
                        break;
                    }
                }
            }

            var nowTime = new DateTimeOffset(DateTime.UtcNow).ToUnixTimeSeconds();
            if (nowTime - startTime > 60)
            {
                break;
            }

            Thread.Sleep(5000);
        }

        client.Disconnect(true);
        return verifyUrl;
    }

    private static string GenerateTelNo()
    {
        var telNo = "512255";
        var rd = new Random();
        for (var i = 0; i < 4; i++)
        {
            telNo += rd.Next(0, 10);
        }

        return telNo;
    }

    private static string GetStateText(string state)
    {
        var stateDictionary = new Dictionary<string, string>
        {
            { "AL", "Alabama" },
            { "AK", "Alaska" },
            { "AZ", "Arizona" },
            { "AR", "Arkansas" },
            { "CA", "California" },
            { "CO", "Colorado" },
            { "CT", "Connecticut" },
            { "DE", "Delaware" },
            { "DC", "District of Columbia" },
            { "FL", "Florida" },
            { "GA", "Georgia" },
            { "HI", "Hawaii" },
            { "ID", "Idaho" },
            { "IL", "Illinois" },
            { "IN", "Indiana" },
            { "IA", "Iowa" },
            { "KS", "Kansas" },
            { "KY", "Kentucky" },
            { "LA", "Louisiana" },
            { "ME", "Maine" },
            { "MD", "Maryland" },
            { "MA", "Massachusetts" },
            { "MI", "Michigan" },
            { "MN", "Minnesota" },
            { "MS", "Mississippi" },
            { "MO", "Missouri" },
            { "MT", "Montana" },
            { "NE", "Nebraska" },
            { "NV", "Nevada" },
            { "NH", "New Hampshire" },
            { "NJ", "New Jersey" },
            { "NM", "New Mexico" },
            { "NY", "New York" },
            { "NC", "North Carolina" },
            { "ND", "North Dakota" },
            { "OH", "Ohio" },
            { "OK", "Oklahoma" },
            { "OR", "Oregon" },
            { "PA", "Pennsylvania" },
            { "RI", "Rhode Island" },
            { "SC", "South Carolina" },
            { "SD", "South Dakota" },
            { "TN", "Tennessee" },
            { "TX", "Texas" },
            { "UT", "Utah" },
            { "VT", "Vermont" },
            { "VA", "Virginia" },
            { "WA", "Washington" },
            { "WV", "West Virginia" },
            { "WI", "Wisconsin" },
            { "WY", "Wyoming" }
        };
        return stateDictionary[state];
    }

    private static string GetMonthText(string month)
    {
        var monthDictionary = new Dictionary<string, string>
        {
            { "1", "January" },
            { "2", "February" },
            { "3", "March" },
            { "4", "April" },
            { "5", "May" },
            { "6", "June" },
            { "7", "July" },
            { "8", "August" },
            { "9", "September" },
            { "10", "October" },
            { "11", "November" },
            { "12", "December" }
        };
        return monthDictionary[month];
    }

    private async void RunPlaywright()
    {
        outputTextBox.Clear();
        var data = CsvToDataTable(Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + @"\peyReg\data.csv",
            1);

        var ti = new CultureInfo("en-US", false).TextInfo;
        for (int i = 0; i < data.Rows.Count; i++)
        {
            var number = data.Rows[i]["No"].ToString();
            var emailInfo = data.Rows[i]["Email"].ToString().Split("----");
            var bankRouter = data.Rows[i]["Router"].ToString();
            var bankAcc = data.Rows[i]["BankAcc"].ToString();
            var password = $"Huobi{data.Rows[i]["No"]}!!";
            var firstName = data.Rows[i]["FirstName"].ToString();
            var lastName = data.Rows[i]["LastName"].ToString();
            var address = ti.ToTitleCase(ti.ToLower(data.Rows[i]["Address"].ToString()));
            var city = ti.ToTitleCase(ti.ToLower(data.Rows[i]["City"].ToString()));
            var state = data.Rows[i]["State"].ToString();
            var postCode = data.Rows[i]["PostCode"].ToString();
            if (postCode.Length == 4)
            {
                postCode = "0" + postCode;
            }

            var ssn = data.Rows[i]["SSN"].ToString();
            if (ssn.Length == 8)
            {
                ssn = "0" + ssn;
            }

            var birthday = data.Rows[i]["Birthday"].ToString().Split("-");
            var telephone = GenerateTelNo();

            // 如果911连接失败，在这里设置一个变量控制是否连接，用while循环来判断，911流程写入循环里

            // 开启911
            is911Applied = false;
            var proxyPort = Apply911Proxy();
            while (!is911Applied)
            {
                Thread.Sleep(1000);
            }

            var apply911Info = apply911String.Split("|");
            if (apply911Info[0] == "failed")
            {
                outputTextBox.AppendText($"Open 911 proxy failed: {apply911Info[1]}" + Environment.NewLine);
                outputTextBox.AppendText("The process will stop." + Environment.NewLine);
                break;
            }

            // 输出911连接信息
            outputTextBox.AppendText($"Using 911 proxy on port: {proxyPort}" + Environment.NewLine);
            outputTextBox.AppendText($"IP: {apply911Info[1]}" + Environment.NewLine);
            outputTextBox.AppendText($"Ping: {apply911Info[2]}" + Environment.NewLine);
            outputTextBox.AppendText($"Country: {apply911Info[3]}" + Environment.NewLine);
            outputTextBox.AppendText($"State: {apply911Info[4]}" + Environment.NewLine);
            outputTextBox.AppendText($"City: {apply911Info[5]}" + Environment.NewLine);
            outputTextBox.AppendText($"Proxies left: {apply911Info[6]}" + Environment.NewLine);

            // 用911socks5开始注册流程
            var playwright = await Playwright.CreateAsync();
            var context = await playwright.Chromium.LaunchPersistentContextAsync(
                Environment.GetFolderPath(Environment.SpecialFolder.Desktop) +
                @$"\peyReg\cache\{number}\", new BrowserTypeLaunchPersistentContextOptions
                {
                    Channel = "chrome",
                    Headless = false,
                    Locale = "en-US",
                    Proxy = new Proxy
                    {
                        Server = $"socks://127.0.0.1:{proxyPort}",
                    },
                    Timeout = 0,
                });
            context.SetDefaultTimeout(0);
            try
            {
                var page = await context.NewPageAsync();
                await page.GotoAsync("https://peymynt.com");
                await page.WaitForLoadStateAsync(LoadState.NetworkIdle);

                // 填写注册信息
                try
                {
                    await page.Locator("[placeholder=\"Enter Your Email Address\"]")
                        .FillAsync(emailInfo[0]);
                    await page.Locator("text=Sign Up with Peymynt").ClickAsync();
                    await page.WaitForLoadStateAsync(LoadState.NetworkIdle);
                }
                catch (Exception e)
                {
                    outputTextBox.AppendText(e.Message + Environment.NewLine);
                }

                // 填写详细注册信息并防止Recaptcha错误
                // 5分钟未成功注册抛出超时异常给顶部catch
                var regStartTime = new DateTimeOffset(DateTime.UtcNow).ToUnixTimeSeconds();
                while (true)
                {
                    await page.Locator("[placeholder=\"First name\"]")
                        .FillAsync(firstName);
                    await page.Locator("[placeholder=\"Last name\"]")
                        .FillAsync(lastName);
                    await page.Locator("[placeholder=\"Password\"]")
                        .FillAsync(password);
                    await page.Locator("[placeholder=\"Confirm password\"]")
                        .FillAsync(password);
                    await page.Locator("button:has-text(\"Sign Up\")").ClickAsync();
                    await page.WaitForLoadStateAsync(LoadState.NetworkIdle);
                    try
                    {
                        await page.Locator("text=Email Verification")
                            .WaitForAsync(new LocatorWaitForOptions { Timeout = 15000 });
                        break;
                    }
                    catch (Exception e)
                    {
                        var regNowTime = new DateTimeOffset(DateTime.UtcNow).ToUnixTimeSeconds();
                        if (regNowTime - regStartTime > 300)
                        {
                            throw new Exception("Can not resolve recaptcha error!");
                        }

                        await page.ReloadAsync();
                        await page.WaitForLoadStateAsync(LoadState.NetworkIdle);
                    }
                }

                // 注册成功密码写入CSV
                data.Rows[i]["Password"] = password;
                DataTaleToCsv(data,
                    Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + @"\peyReg\data.csv");

                // 检查邮箱
                var verifyUrl = "";
                while (true)
                {
                    verifyUrl = CheckEmailForURL(emailInfo);
                    if (verifyUrl != "")
                    {
                        break;
                    }

                    await page.Locator("text=click here to resend.").ClickAsync();
                    await page.Locator("[placeholder=\"Enater Resend Email address\"]")
                        .FillAsync(emailInfo[0]);
                    await page.Locator("button:has-text(\"Resend Email\")").ClickAsync();
                    await page.WaitForLoadStateAsync(LoadState.NetworkIdle);
                }

                await page.CloseAsync();

                var confirmPage = await context.NewPageAsync();
                await confirmPage.GotoAsync(verifyUrl);
                await confirmPage.WaitForLoadStateAsync(LoadState.NetworkIdle);

                // 登录并防止Recaptcha错误
                await confirmPage.Locator("[placeholder=\"Email address\"]").FillAsync(emailInfo[0]);
                await confirmPage.Locator("[placeholder=\"Password\"]").FillAsync(password);
                while (true)
                {
                    await confirmPage.Locator("button:has-text(\"Sign In\")").ClickAsync();
                    await confirmPage.WaitForLoadStateAsync(LoadState.NetworkIdle);
                    try
                    {
                        await confirmPage.Locator("[placeholder=\"Company name\"]")
                            .WaitForAsync(new LocatorWaitForOptions { Timeout = 15000 });
                        break;
                    }
                    catch (Exception e)
                    {
                    }
                }

                // 填写业务基本信息
                try
                {
                    await confirmPage.Locator("[placeholder=\"Company name\"]")
                        .FillAsync($"{firstName} {lastName}");
                    await confirmPage.Locator("text=Type of businessType of businessThis field is required >> span")
                        .Nth(1)
                        .ClickAsync();
                    await confirmPage.Locator("text=Artists, Photographers & Creative Types").ClickAsync();
                    await confirmPage.Locator("text=SubtypeActorThis field is required >> span").Nth(2).ClickAsync();
                    await confirmPage.Locator("[aria-label=\"Photographer\"]").ClickAsync();
                    // 看看是否未出现国家选项
                    try
                    {
                        await confirmPage.Locator("text=Change this")
                            .ClickAsync(new LocatorClickOptions { Timeout = 1000 });
                    }
                    catch (Exception e)
                    {
                        await confirmPage.Locator("text=CountryCountryThis field is required >> span").Nth(1)
                            .ClickAsync();
                        await confirmPage.Locator("text=United States").ClickAsync();
                    }

                    await confirmPage.Locator("text=Type of entityType of entityThis field is required >> span").Nth(1)
                        .ClickAsync();
                    await confirmPage.Locator("text=Individuals").ClickAsync();
                    await confirmPage.Locator("text=Get started").ClickAsync();
                    await confirmPage.WaitForLoadStateAsync(LoadState.NetworkIdle);
                }
                catch (Exception e)
                {
                    outputTextBox.AppendText(e.Message + Environment.NewLine);
                }

                // 设置订阅Starter计划
                try
                {
                    await confirmPage.Locator(".py-form__element__faux").First.ClickAsync();
                    await confirmPage.Locator("text=Get Started").ClickAsync();
                    await confirmPage.WaitForLoadStateAsync(LoadState.NetworkIdle);
                }
                catch (Exception e)
                {
                    outputTextBox.AppendText(e.Message + Environment.NewLine);
                }

                // 在Dashboard进入激活流程
                try
                {
                    await confirmPage.Locator("text=Activate Payments Now").ClickAsync();
                    await confirmPage.WaitForLoadStateAsync(LoadState.NetworkIdle);
                }
                catch (Exception e)
                {
                    outputTextBox.AppendText(e.Message + Environment.NewLine);
                }

                // 点击激活按钮
                try
                {
                    await confirmPage
                        .Locator(
                            "text=Payments by PeymyntYour customers can pay you online.Get paid by your customers  >> button")
                        .ClickAsync();
                    await confirmPage.WaitForLoadStateAsync(LoadState.NetworkIdle);
                }
                catch (Exception e)
                {
                    outputTextBox.AppendText(e.Message + Environment.NewLine);
                }

                // 选择业务类型为个人
                try
                {
                    await confirmPage.Locator("text=Individuals and Sole Proprietorships").ClickAsync();
                    await confirmPage.Locator("text=Save and continue").ClickAsync();
                    await confirmPage.WaitForLoadStateAsync(LoadState.NetworkIdle);
                }
                catch (Exception e)
                {
                    outputTextBox.AppendText(e.Message + Environment.NewLine);
                }

                // 填写业务详细信息
                try
                {
                    await confirmPage.Locator("[placeholder=\"Business Legal Name\"]")
                        .FillAsync($"{firstName} {lastName}");
                    await confirmPage.Locator("text=Select a category").ClickAsync();
                    await confirmPage.Locator("text=Photographic Studios").ClickAsync();
                    await confirmPage.Locator("label:has-text(\"Products\")").ClickAsync();
                    await confirmPage.Locator("textarea[name=\"description\"]")
                        .FillAsync($"{firstName} {lastName}");
                    await confirmPage.Locator("text=Telephone+This field is required >> input[type=\"text\"]")
                        .FillAsync(telephone);
                    await confirmPage.Locator("[placeholder=\"Street\"]")
                        .FillAsync(address);
                    await confirmPage.WaitForTimeoutAsync(3000);
                    await confirmPage.Locator("[placeholder=\"www\\.example\\.com\"]").ClickAsync();
                    await confirmPage.WaitForTimeoutAsync(1000);
                    await confirmPage.Locator("[placeholder=\"City\"]").FillAsync(city);
                    await confirmPage.Locator("text=StateStatethis field is required >> span").Nth(1).ClickAsync();
                    await confirmPage.Locator($"div.Select-menu-outer :text-is(\"{GetStateText(state)}\")")
                        .ClickAsync();
                    await confirmPage.Locator("[placeholder=\"ZIP\"]").FillAsync(postCode);
                    await confirmPage.Locator("text=Have you accepted credit card payments in the past?YesNo >> span")
                        .First
                        .ClickAsync();
                    await confirmPage.Locator("text=Save and continue").ClickAsync();
                    await confirmPage.WaitForLoadStateAsync(LoadState.NetworkIdle);
                }
                catch (Exception e)
                {
                    outputTextBox.AppendText(e.Message + Environment.NewLine);
                }

                // 填写个人信息（SSN）
                try
                {
                    await confirmPage.Locator("[placeholder=\"Legal First Name\"]").FillAsync(firstName);
                    await confirmPage.Locator("[placeholder=\"Legal Last Name\"]").FillAsync(lastName);
                    await confirmPage.Locator("[placeholder=\"\\39 99-99-9999\"]").FillAsync(ssn);
                    await confirmPage.Locator("text=Month").ClickAsync();
                    await confirmPage.Locator($"text={GetMonthText(birthday[0])}").ClickAsync();
                    await confirmPage.Locator("text=Day").ClickAsync();
                    if (birthday[1].Length == 1)
                    {
                        await confirmPage.Locator($"[aria-label=\"\\3{birthday[1]} \"]").ClickAsync();
                    }
                    else if (birthday[1].Length == 2)
                    {
                        await confirmPage.Locator($"div.Select-menu-outer :text-is(\"{birthday[1]}\")").ClickAsync();
                    }

                    await confirmPage.Locator("[placeholder=\"Year\"]").FillAsync(birthday[2]);
                    await confirmPage.Locator("input[type=\"number\"]").FillAsync(telephone);
                    await confirmPage.Locator("[placeholder=\"Email\"]").FillAsync(emailInfo[0]);
                    await confirmPage.Locator("text=Save and continue").ClickAsync();
                    await confirmPage.WaitForLoadStateAsync(LoadState.NetworkIdle);
                }
                catch (Exception e)
                {
                    outputTextBox.AppendText(e.Message + Environment.NewLine);
                }

                // 填写银行路由和账号
                try
                {
                    await confirmPage.Locator("text=Manually connect bank").ClickAsync();
                    await confirmPage.Locator("[placeholder=\"Routing Number\"]").FillAsync(bankRouter);
                    await confirmPage.Locator("[placeholder=\"Account Number\"]").FillAsync(bankAcc);
                    await confirmPage.Locator("text=Save and continue").ClickAsync();
                    await confirmPage.WaitForLoadStateAsync(LoadState.NetworkIdle);
                }
                catch (Exception e)
                {
                    outputTextBox.AppendText(e.Message + Environment.NewLine);
                }

                // 同意协议并确认
                try
                {
                    await confirmPage.Locator(".py-form__element__faux").ClickAsync();
                    await confirmPage.Locator("text=Save and continue").ClickAsync();
                    await confirmPage.WaitForLoadStateAsync(LoadState.NetworkIdle);
                    await confirmPage.Locator("h2:has-text(\"Payments\")")
                        .WaitForAsync(new LocatorWaitForOptions { Timeout = 300000 });
                }
                catch (Exception e)
                {
                    // 5分钟还没过KYC
                    // 先检查是不是有生成发票样式的流程
                    var isInvoiceVisible = true;
                    try
                    {
                        await confirmPage.Locator("text=Create a new invoice")
                            .ClickAsync(new LocatorClickOptions { Timeout = 5000 });
                        await confirmPage.Locator("text=Looks good, let's go").ClickAsync();
                        await confirmPage.Locator("h2:has-text(\"Invoices\")").WaitForAsync();
                    }
                    catch (Exception ee)
                    {
                        isInvoiceVisible = false;
                    }

                    if (!isInvoiceVisible)
                    {
                        data.Rows[i]["Link"] = "Invalid";
                        DataTaleToCsv(data,
                            Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + @"\peyReg\data.csv");

                        await confirmPage.CloseAsync();
                        await context.CloseAsync();
                        playwright.Dispose();

                        outputTextBox.AppendText(
                            "----------------------------------------------------------------------------------------------------------------------------------------------------" +
                            Environment.NewLine);
                        continue;
                    }
                }

                // 获取peylink
                var peyLink = "";
                try
                {
                    await confirmPage.Locator("text=My PeyMe Lynk").ClickAsync();
                    await confirmPage.WaitForLoadStateAsync(LoadState.NetworkIdle);
                    await confirmPage.Locator("text=Create PeyMe Lynk now").ClickAsync();
                    await confirmPage.Locator("input[name=\"userName\"]")
                        .FillAsync($"{firstName.Replace(" ", "")}{lastName.Replace(" ", "")}{telephone.Substring(6)}");
                    await confirmPage.Locator("text=Confirm").ClickAsync();
                    await confirmPage.Locator(
                            "text=I agree to the Terms & Conditions and Privacy Policy of the Peymynt platform.Rea >> span")
                        .First.ClickAsync();
                    await confirmPage.Locator("text=Ready to use").ClickAsync();
                    await confirmPage.WaitForLoadStateAsync(LoadState.NetworkIdle);
                    await confirmPage.Locator("text=Done").ClickAsync();
                    await confirmPage.WaitForLoadStateAsync(LoadState.NetworkIdle);
                    peyLink = await confirmPage.Locator(".py-share-link a").TextContentAsync();
                    outputTextBox.AppendText(peyLink + Environment.NewLine);
                }
                catch (Exception e)
                {
                    outputTextBox.AppendText(e.Message + Environment.NewLine);
                }

                data.Rows[i]["Link"] = peyLink;
                DataTaleToCsv(data,
                    Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + @"\peyReg\data.csv");

                await confirmPage.CloseAsync();
                await context.CloseAsync();
                playwright.Dispose();
            }
            catch (Exception ge)
            {
                outputTextBox.AppendText(ge.Message + Environment.NewLine);
                await context.CloseAsync();
                playwright.Dispose();
            }

            outputTextBox.AppendText(
                "----------------------------------------------------------------------------------------------------------------------------------------------------" +
                Environment.NewLine);
        }

        startButton.Enabled = true;
    }

    private void startButton_Click(object sender, EventArgs e)
    {
        startButton.Enabled = false;
        Task.Run(RunPlaywright);
    }
}