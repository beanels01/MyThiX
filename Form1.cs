using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using System.IO;
using System.Threading;




namespace WindowsFormsApp1
{


    public partial class Form1 : Form
    {
        String VersionCode = "v1.4.1";
        String PublicTotalForm = "1AxlF5nYUEo4E8oJw_x2hMKicOYYEoefshPAdknIGg40";
        String ServiceTotalForm = "1mz2gaIKqHasRumcF3LfY6tIRGScMRfkYiAQX1LKGy6k";
        String PlayerForm = "1gg1M9Ldrr-YQBRkDJNsADRd1lJxia8X_Gx_dwUGW5aY";
        String ID_Var;
        // If modifying these scopes, delete your previously saved credentials
        // at ~/.credentials\\sheets.googleapis.com-dotnet-quickstart.json
        static string[] Scopes = { SheetsService.Scope.Spreadsheets };
        static string ApplicationName = "Google Sheets API .NET Quickstart";
        public Form1()
        {
            InitializeComponent();
        }


        private void Form1_Load(object sender, EventArgs e)
        {
            string path = "GDID.ini";
            if (!File.Exists(path))
            {
                textBox1.Text = "";
            }
            else
            {
                string readText = File.ReadAllText(path);
                if(readText!="") checkBox1.Checked=true;
                textBox1.Text = readText;
            }

        }

        public int prev_index=-1;
        public int prev_indexPlayer=-1;
        Global Main_Form = new Global();
        Global Skill_Form = new Global();
        Global Service_Form = new Global();
        public string credPath;
        public static int balance_begin = 26;
        public static int balance_end = 33;
        public static int transform_begin = 2;
        public static int transform_end = 7;
        public static int practice_red_top_begin = 58;
        public static int practice_red_bottom_begin = 69;
        public static int practice_red_hand_begin = 87;
        public static int practice_red_shoes_begin = 107;
        public static int practice_red_end = 124;
        public static int practice_blue_top_begin = 182;
        public static int practice_blue_bottom_begin = 193;
        public static int practice_blue_hand_begin = 211;
        public static int practice_blue_shoes_begin = 231;
        public static int practice_blue_end = 248;
        public static int practice_purple_top_begin = 306;
        public static int practice_purple_bottom_begin = 317;
        public static int practice_purple_hand_begin = 338;
        public static int practice_purple_shoes_begin = 358;
        public static int practice_purple_end = 375;
        public static int skill_red_begin = 3;
        public static int skill_blue_begin = 203;
        public static int skill_purple_begin = 475;
        public static int skill_purple_end = 910;
        public UserCredential credential;


        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void label3_Click(object sender, EventArgs e)
        {

        }

        private void label6_Click(object sender, EventArgs e)
        {

        }


        public class Global
        {
            public IList<IList<Object>> g_Form;
        }

        public class moneyitem
        {
            public string player;
            public string color;
            public string kind;
            public string ability;
            public string equip;
            public string per;
            public string value { get; set; }
            public int price { get; set; }
        }

        public void UploadTransaction(moneyitem mi)
        {

        }
        public void GenerateItemList(Global Main_Form, Global Skill_Form, ComboBox cc, ComboBox ce, ComboBox ck, ComboBox ca)
        {
            
            if (ck.Text.ToString() == "均衡")
            {

                ca.Items.Clear();
                for (int i = balance_begin; i <= balance_end; i++)
                    try
                    {
                        ca.Items.Add(Main_Form.g_Form[i][2]);
                    }
                    catch (Exception e)
                    {
                        continue;
                    }
                ca.SelectedIndex = 0;
            }

            else if (ck.Text.ToString() == "變化")
            {
                ca.Items.Clear();
                for (int i = transform_begin; i <= transform_end; i++)
                    try
                    {
                        ca.Items.Add(Main_Form.g_Form[i][2]);
                    }
                    catch (Exception e)
                    {
                        continue;
                    }
                ca.SelectedIndex = 0;

            }

            else if (ck.Text.ToString() == "熟練")
            {
                int range_begin = 0, range_end = 0;
                if (cc.Text.ToString() == "赤紅")
                {
                    if (ce.Text.ToString() == "上衣")
                    {
                        range_begin = practice_red_top_begin;
                        range_end = practice_red_bottom_begin - 1;
                    }
                    else if (ce.Text.ToString() == "下衣")
                    {
                        range_begin = practice_red_bottom_begin;
                        range_end = practice_red_hand_begin - 1;
                    }
                    else if (ce.Text.ToString() == "手套")
                    {
                        range_begin = practice_red_hand_begin;
                        range_end = practice_red_shoes_begin - 1;
                    }
                    else if (ce.Text.ToString() == "鞋子")
                    {
                        range_begin = practice_red_shoes_begin;
                        range_end = practice_red_end;
                    }
                }
                else if (cc.Text.ToString() == "青藍")
                {
                    if (ce.Text.ToString() == "上衣")
                    {
                        range_begin = practice_blue_top_begin;
                        range_end = practice_blue_bottom_begin - 1;
                    }
                    else if (ce.Text.ToString() == "下衣")
                    {
                        range_begin = practice_blue_bottom_begin;
                        range_end = practice_blue_hand_begin - 1;
                    }
                    else if (ce.Text.ToString() == "手套")
                    {
                        range_begin = practice_blue_hand_begin;
                        range_end = practice_blue_shoes_begin - 1;
                    }
                    else if (ce.Text.ToString() == "鞋子")
                    {
                        range_begin = practice_blue_shoes_begin;
                        range_end = practice_blue_end;
                    }
                }
                else if (cc.Text.ToString() == "靛紫")
                {
                    if (ce.Text.ToString() == "上衣")
                    {
                        range_begin = practice_purple_top_begin;
                        range_end = practice_purple_bottom_begin - 1;
                    }
                    else if (ce.Text.ToString() == "下衣")
                    {
                        range_begin = practice_purple_bottom_begin;
                        range_end = practice_purple_hand_begin - 1;
                    }
                    else if (ce.Text.ToString() == "手套")
                    {
                        range_begin = practice_purple_hand_begin;
                        range_end = practice_purple_shoes_begin - 1;
                    }
                    else if (ce.Text.ToString() == "鞋子")
                    {
                        range_begin = practice_purple_shoes_begin;
                        range_end = practice_purple_end;
                    }
                }

                ca.Items.Clear();
                for (int i = range_begin; i <= range_end; i++)
                    try
                    {
                        ca.Items.Add(Main_Form.g_Form[i][2]);
                        Console.WriteLine("{0}", Main_Form.g_Form[i][2]);
                    }
                    catch (Exception e)
                    {
                        continue;
                    }
                ca.SelectedIndex = 0;
                Console.Read();

            }
            else if (ck.Text.ToString() == "破壞")
            {
                int range_begin = 0, range_end = 0;
                if (cc.Text.ToString() == "赤紅")
                {
                    range_begin = skill_red_begin;
                    range_end = skill_blue_begin - 1;
                }
                if (cc.Text.ToString() == "青藍")
                {
                    range_begin = skill_blue_begin;
                    range_end = skill_purple_begin - 1;
                }
                if (cc.Text.ToString() == "靛紫")
                {
                    range_begin = skill_purple_begin;
                    range_end = skill_purple_end;
                }
                ca.Items.Clear();
                for (int i = range_begin; i <= range_end; i++)
                    try
                    {
                        if (Skill_Form.g_Form[i][1].ToString() != "")
                        {
                            ca.Items.Add(Skill_Form.g_Form[i][1]);
                            Console.WriteLine("{0}", Main_Form.g_Form[i][1]);
                        }
                    }
                    catch (Exception e)
                    {
                        continue;
                    }
                ca.SelectedIndex = 0;
                Console.Read();
            }

            

        }
        public void calexchange()
        {
            if (comboReq.Text == "預支")
            {
                
                int exval =Convert.ToInt32(comboAmount.Text);
                exval *= 22;
                exval /= 100;
                exchangeamount.Text = exval.ToString();

            }
            else
            {
                int value = comboEx.SelectedIndex + 1;
                if (value > 2)
                {
                    int exval = 65 * Convert.ToInt32(textBox3.Text);
                    exchangeamount.Text =exval.ToString();
                }
                else {

                    int exval = 2 * Convert.ToInt32(textBox3.Text);
                    exchangeamount.Text = exval.ToString();
                }
            }
        }

        public void Login_Success()
        {
            string OnlineVersionCode="";
            var service = new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });
            for (int i = 0; i < Main_Form.g_Form.Count; i++)
            {

                    if(i!=249 && i!=125 && i!=1)
                    try
                    {
                            if (Service_Form.g_Form[i][4].ToString() != Main_Form.g_Form[i][3].ToString())
                            {
                                status.Text = "目前狀態：偵測到價格異動...同步中";
                                Console.WriteLine("偵測到價格異動 patching... x:" + i);
                                Main_Form.g_Form[i][3] = Service_Form.g_Form[i][4];
                                string range = "非技能類!D" + (i+1).ToString();
                                ValueRange valueRange = new ValueRange();
                                valueRange.MajorDimension = "ROWS";
                                var oblist = new List<object>() { Main_Form.g_Form[i][3] };
                                valueRange.Values = new List<IList<object>> { oblist };
                                Console.WriteLine("Uploading Array...");
                                SpreadsheetsResource.ValuesResource.UpdateRequest update = service.Spreadsheets.Values.Update(valueRange, ID_Var, range);
                                update.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.RAW;
                                UpdateValuesResponse result2 = update.Execute();
                    }
                    }
                    catch(Exception er)
                    {
                        MessageBox.Show(er.ToString());
                        continue;
                    }
            }
            for(int i = Main_Form.g_Form.Count; i < Service_Form.g_Form.Count; i++)
            {
                try
                {
                    if (Service_Form.g_Form[i][0].ToString() == "後端版本號")
                    {
                         OnlineVersionCode = Service_Form.g_Form[i][1].ToString();
                    }
                }
                catch
                {
                    continue;
                }
            }
            


            if (VersionCode.StartsWith(OnlineVersionCode))
            {
                for (int i = 5; i < Main_Form.g_Form[1].Count; i++)
                {
                    if (Main_Form.g_Form[1][i].ToString() != "")
                    {
                        ComboPlayer.Items.Add(Main_Form.g_Form[1][i]);
                        ComboPlayer2.Items.Add(Main_Form.g_Form[1][i]);
                        comboApplier.Items.Add(Main_Form.g_Form[1][i]);
                    }
                    else
                        break;
                }
                status.Text = "目前狀態:登入成功!歡迎使用小助手" + VersionCode + "!";
                AutoFixForm();
                tabControl1.Enabled = true;
                ComboPlayer.SelectedIndex = 0;
                ComboPlayer2.SelectedIndex = 0;
                comboReq.SelectedIndex = 0;
                comboApplier.SelectedIndex = 0;
                comboEx.SelectedIndex = 0;
                GenerateItemList(Main_Form, Skill_Form, ComboColor, ComboEquip, ComboKind, ComboAbility);
                GenerateItemList(Main_Form, Skill_Form, ComboColor2, ComboEquip2, ComboKind2, ComboAbility2);
            }
            else
            {
                Console.WriteLine(VersionCode);
                MessageBox.Show("小助手版本過舊！請向幹部索取更新版本！");
            }
        }


        

        private void ComboKind_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (ComboKind.Text.ToString() == "破壞")
            {
                ComboPer.Enabled = true;
                ComboPer.SelectedIndex = 0;
            }
            else
            {
                ComboPer.Enabled = false;
            }
           GenerateItemList(Main_Form,Skill_Form,ComboColor,ComboEquip,ComboKind,ComboAbility);

        }

        private void ComboColor_SelectedIndexChanged(object sender, EventArgs e)
        {
            GenerateItemList(Main_Form, Skill_Form, ComboColor, ComboEquip, ComboKind, ComboAbility);
        }

        private void comboEquip_SelectedIndexChanged(object sender, EventArgs e)
        {
            GenerateItemList(Main_Form, Skill_Form, ComboColor, ComboEquip, ComboKind, ComboAbility);

        }

        private void ComboKind_QueryAccessibilityHelp(object sender, QueryAccessibilityHelpEventArgs e)
        {

        }

        private void ComboAbility_SelectedIndexChanged(object sender, EventArgs e)
        {
            int itemIndex=0;
            if (ComboKind.Text.ToString() == "破壞")
            {

                for (int i = 0; i < Skill_Form.g_Form.Count; i++)
                {
                    try
                    {
                        if (Skill_Form.g_Form[i][1].ToString() == ComboAbility.Text.ToString())
                        {
                            itemIndex = i;
                            break;
                        }
                    }
                    catch(Exception exr)
                    {
                        continue;
                    }
                }
                price.Text = Skill_Form.g_Form[itemIndex][2].ToString();
            }
            else
            {
                for (int i = 0; i < Main_Form.g_Form.Count; i++)
                {
                    try
                    {
                        if (Main_Form.g_Form[i][2].ToString() == ComboAbility.Text.ToString())
                        {
                            itemIndex = i;
                            break;
                        }
                    }
                    catch(Exception exp)
                    {
                        continue;
                    }
                }
                price.Text = Main_Form.g_Form[itemIndex][3].ToString();
            }
        }
        public void AutoFixForm()
        {

        }


        private void tabPage1_Click(object sender, EventArgs e)
        {

        }




        private void ComboValue_SelectedIndexChanged(object sender, EventArgs e)
        {
        }

        private void label4_Click(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (textBox2.Text != "988386")
            {
                foreach (Control ctl in tabPage4.Controls) ctl.Enabled = false;
            }

            if(checkBox1.Checked)
            {
                string path = "GDID.ini";
                if (!File.Exists(path))
                {
                    string createText = textBox1.Text;
                    File.WriteAllText(path, createText);
                }
                else
                {
                    string createText = textBox1.Text;
                    File.WriteAllText(path, createText);
                }
                
            }
            else
            {
                string path = "GDID.ini";
                    string createText = "";
                File.WriteAllText(path, createText);
            }
            try
            {
               

                using (var stream =
                    new FileStream("client_secret.json", FileMode.Open, FileAccess.Read))
                {
                        credPath = ".credentials\\sheets.googleapis.com-dotnet-quickstart.json";
                        credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                            GoogleClientSecrets.Load(stream).Secrets,
                            Scopes,
                            "user",
                            CancellationToken.None,
                            new FileDataStore(credPath, true)).Result;
                        Console.WriteLine("Credential file saved to: " + credPath);
                }
                // Create Google Sheets API service.
                var service = new SheetsService(new BaseClientService.Initializer()
                {
                    HttpClientInitializer = credential,
                    ApplicationName = ApplicationName,
                });
                status.Text = "目前狀態:登入成功!請稍候系統自動同步表單訊息....";

                ID_Var = textBox1.Text;
                SpreadsheetsResource.ValuesResource.GetRequest request =
                service.Spreadsheets.Values.Get(ID_Var,"非技能類!A1:P523");
                ValueRange response = request.Execute();
                Main_Form.g_Form = response.Values;
                request =
                service.Spreadsheets.Values.Get(ID_Var, "技能類!A1:BL911");
                response = request.Execute();
                Skill_Form.g_Form = response.Values;
                request =
                service.Spreadsheets.Values.Get(ServiceTotalForm, "非技能類!A1:F378");
                response = request.Execute();
                Service_Form.g_Form = response.Values;
                Login_Success();
            }
            catch(Exception ex_login) {
                MessageBox.Show("表單取得錯誤!請修正試算表ID\n錯誤訊息："+ex_login.ToString());
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (ComboAbility.Text.ToString() == "" || ComboPer.Text.ToString()=="")
            {
                status.Text = "請正確填寫所有碎片資訊!";
                status.ForeColor = System.Drawing.Color.Red;
            }
            else
            {
                UserCredential credential;
                using (var stream =
                    new FileStream("client_secret.json", FileMode.Open, FileAccess.Read))
                {

                    credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                        GoogleClientSecrets.Load(stream).Secrets,
                        Scopes,
                        "user",
                        CancellationToken.None,
                        new FileDataStore(credPath, true)).Result;
                    Console.WriteLine("Credential file saved to: " + credPath);
                }

                int itemIndex = -1;
                int itemIndexPlayer = -1;

                if (ComboKind.Text.ToString() == "破壞")
                {
                    for (int i = 0; i < Skill_Form.g_Form.Count; i++)
                    {
                        if (Skill_Form.g_Form[i][0].ToString() == ComboColor.Text.ToString())
                        {
                            itemIndex = i;
                            break;
                        }
                    }
                    for (int i = itemIndex; i < Skill_Form.g_Form.Count; i++)
                    {
                        if (Skill_Form.g_Form[i][1].ToString() == ComboAbility.Text.ToString())
                        {
                            itemIndex = i;
                            break;
                        }
                    }
                    for (int i = itemIndex; i < Skill_Form.g_Form.Count; i++)
                    {
                        if (Skill_Form.g_Form[i][2].ToString() == ComboEquip.Text.ToString())
                        {
                            itemIndex = i;
                            break;
                        }
                    }
                    for (int i = 8; i < Skill_Form.g_Form[1].Count; i++)
                    {
                        if (Skill_Form.g_Form[1][i].ToString() == ComboPlayer.Text.ToString())
                        {
                            itemIndexPlayer = i;
                            break;
                        }
                    }
                    for (int i = itemIndexPlayer; i < Skill_Form.g_Form[2].Count; i++)
                    {
                        if (Skill_Form.g_Form[2][i].ToString() == ComboPer.Text.ToString())
                        {
                            itemIndexPlayer = i;
                            break;
                        }
                    }
                    prev_index = itemIndex;
                    prev_indexPlayer = itemIndexPlayer;
                    ///////
                    string parsepost = "";
                    itemIndexPlayer++;
                    itemIndex++;
                    

                    Console.WriteLine("position {0} {1} is now {2}", itemIndex, itemIndexPlayer, Skill_Form.g_Form[itemIndex][itemIndexPlayer]);


                    if (itemIndexPlayer > 26)
                    {
                        parsepost += "A";
                        itemIndexPlayer -= 26;
                    }

                    parsepost += Convert.ToChar(itemIndexPlayer + 'A' - 1);


                    var service = new SheetsService(new BaseClientService.Initializer()
                    {
                        HttpClientInitializer = credential,
                        ApplicationName = ApplicationName,
                    });

                    string range = "技能類!" + parsepost + itemIndex.ToString();
                    Console.WriteLine(range);
                    int stock_num = Convert.ToInt32(Skill_Form.g_Form[itemIndex - 1][itemIndexPlayer - 1]);
                    stock_num++;
                    Console.WriteLine("updated stock is now: " + stock_num.ToString());
                    Skill_Form.g_Form[itemIndex - 1][itemIndexPlayer - 1] = stock_num.ToString();
                    Console.WriteLine("updated skill form is now: " + Skill_Form.g_Form[itemIndex - 1][itemIndexPlayer - 1]);
                    position.Text = range;
                    num_temp.Text = stock_num.ToString();
                    var oblist = new List<object>() { stock_num };
                    ValueRange valueRange = new ValueRange();
                    valueRange.MajorDimension = "ROWS";
                    valueRange.Values = new List<IList<object>> { oblist };
                    Console.WriteLine("value range num is now: " + valueRange.Values[0][0]);
                    SpreadsheetsResource.ValuesResource.UpdateRequest update = service.Spreadsheets.Values.Update(valueRange, ID_Var, range);
                    update.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.RAW;
                    UpdateValuesResponse result2 = update.Execute();

                }
                else
                {
                    for (int i = 0; i < Main_Form.g_Form.Count; i++)
                    {
                        try
                        {
                            if (Main_Form.g_Form[i][0].ToString() == ComboColor.Text.ToString() + "/" + ComboKind.Text.ToString())
                            {
                                itemIndex = i;
                                break;
                            }
                        }
                        catch
                        {
                            continue;
                        }
                    }
                    for (int i = itemIndex; i < Main_Form.g_Form.Count; i++)
                    {
                        try
                        {
                            if (Main_Form.g_Form[i][1].ToString() == ComboEquip.Text.ToString())
                            {
                                itemIndex = i;
                                break;
                            }
                        }
                        catch
                        {
                            continue;
                        }
                    }
                    for (int i = itemIndex; i < Main_Form.g_Form.Count; i++)
                        {
                            try
                            {
                                if (Main_Form.g_Form[i][2].ToString() == ComboAbility.Text.ToString())
                                {
                                    itemIndex = i;
                                    break;
                                }
                            }
                            catch
                            {
                                continue;
                            }
                        
                    }
                    for (int i = 5; i <= Main_Form.g_Form[1].Count; i++)
                    {
                        if (Main_Form.g_Form[1][i].ToString() == ComboPlayer.Text.ToString())
                        {
                            itemIndexPlayer = i;
                            break;
                        }
                    }
                    prev_index = itemIndex;
                    prev_indexPlayer = itemIndexPlayer;


                    //////
                    string parsepost = "";
                    itemIndexPlayer++;
                    itemIndex++;


                    if (itemIndexPlayer > 26)
                    {
                        parsepost += "A";
                        itemIndexPlayer -= 26;
                    }
                    
                    var service = new SheetsService(new BaseClientService.Initializer()
                    {
                        HttpClientInitializer = credential,
                        ApplicationName = ApplicationName,
                    });

                    parsepost += Convert.ToChar(itemIndexPlayer+'A'-1);
                    string range = "非技能類!"+parsepost+itemIndex.ToString();
                    Console.WriteLine(range);

                    int stock_num = Convert.ToInt32(Main_Form.g_Form[itemIndex - 1][itemIndexPlayer - 1]);
                    stock_num++;
                    Main_Form.g_Form[itemIndex - 1][itemIndexPlayer-1] = stock_num.ToString();
                    position.Text = range;
                    num_temp.Text = stock_num.ToString();
                    var oblist = new List<object>() { stock_num };
                    ValueRange valueRange = new ValueRange();
                    valueRange.MajorDimension = "ROWS";
                    valueRange.Values = new List<IList<object>> { oblist };
                    Console.WriteLine("value range num is now: " + valueRange.Values[0][0]);
                    SpreadsheetsResource.ValuesResource.UpdateRequest update = service.Spreadsheets.Values.Update(valueRange, ID_Var, range);
                    update.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.RAW;
                    UpdateValuesResponse result2 = update.Execute();

                }


                status.Text = ComboAbility.Text + "新增成功!請至表單確認!";
                status.ForeColor = System.Drawing.Color.Green;
                button3.Enabled = true;
            }
        }

        private void label5_Click(object sender, EventArgs e)
        {

        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void button3_Click(object sender, EventArgs e)
        {
            DialogResult result = MessageBox.Show("確定要復原嗎？", "不要玩系統！",
            MessageBoxButtons.OKCancel, MessageBoxIcon.Warning);
            if (result == DialogResult.OK)
            {
                UserCredential credential;
                using (var stream =
                    new FileStream("client_secret.json", FileMode.Open, FileAccess.Read))
                {

                    credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                        GoogleClientSecrets.Load(stream).Secrets,
                        Scopes,
                        "user",
                        CancellationToken.None,
                        new FileDataStore(credPath, true)).Result;
                    Console.WriteLine("Credential file saved to: " + credPath);
                }
                // Create Google Sheets API service.
                var service = new SheetsService(new BaseClientService.Initializer()
                {
                    HttpClientInitializer = credential,
                    ApplicationName = ApplicationName,
                });
                int stock_num = Convert.ToInt32(num_temp.Text);
                stock_num--;
                num_temp.Text = stock_num.ToString();
                if (position.Text.StartsWith("非技能類"))
                {
                    int x= Convert.ToInt32(Main_Form.g_Form[prev_index][prev_indexPlayer])-1;
                    Main_Form.g_Form[prev_index][prev_indexPlayer] = x.ToString();
                }
                else
                {
                    int x = Convert.ToInt32(Skill_Form.g_Form[prev_index][prev_indexPlayer])-1;
                    Skill_Form.g_Form[prev_index][prev_indexPlayer] = x.ToString();
                }
                Console.WriteLine(stock_num);
                var oblist = new List<object>() { stock_num };
                ValueRange valueRange = new ValueRange();
                valueRange.MajorDimension = "ROWS";
                valueRange.Values = new List<IList<object>> { oblist };
                Console.WriteLine(valueRange.Values[0][0]);
                SpreadsheetsResource.ValuesResource.UpdateRequest update = service.Spreadsheets.Values.Update(valueRange, ID_Var, position.Text);
                update.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.RAW;
                UpdateValuesResponse result2 = update.Execute();
                status.Text = "復原成功！";
                num_temp.Text = stock_num.ToString();
                status.ForeColor = System.Drawing.Color.DarkGoldenrod;
                button3.Enabled = false;
                }



        }





        /// <summary>
        /// ///////////////////////////////////////////////////////////////////這裡開始是tab2
        /// </summary>
        /// 
        private void ComboKind2_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (ComboKind2.Text.ToString() == "破壞")
            {
                ComboPer2.Enabled = true;
                ComboPer2.SelectedIndex = 0;
            }
            else
            {
                ComboPer2.Enabled = false;
            }
            GenerateItemList(Main_Form, Skill_Form, ComboColor2, ComboEquip2, ComboKind2, ComboAbility2);

        }

        private void ComboColor2_SelectedIndexChanged(object sender, EventArgs e)
        {
            GenerateItemList(Main_Form, Skill_Form, ComboColor2, ComboEquip2, ComboKind2, ComboAbility2);
        }

        private void ComboEquip2_SelectedIndexChanged(object sender, EventArgs e)
        {
            GenerateItemList(Main_Form, Skill_Form, ComboColor2, ComboEquip2, ComboKind2, ComboAbility2);

        }

        private void ComboKind2_QueryAccessibilityHelp(object sender, QueryAccessibilityHelpEventArgs e)
        {

        }

        private void ComboAbility2_SelectedIndexChanged(object sender, EventArgs e)
        {
            int itemIndex = 0;
            if (ComboKind.Text.ToString() == "破壞")
            {

                for (int i = 0; i < Skill_Form.g_Form.Count; i++)
                {
                    try
                    {
                        if (Skill_Form.g_Form[i][1].ToString() == ComboAbility.Text.ToString())
                        {
                            itemIndex = i;
                            break;
                        }
                    }
                    catch (Exception exr)
                    {
                        continue;
                    }
                }
                price.Text = Skill_Form.g_Form[itemIndex][2].ToString();
            }
            else
            {
                for (int i = 0; i < Main_Form.g_Form.Count; i++)
                {
                    try
                    {
                        if (Main_Form.g_Form[i][2].ToString() == ComboAbility.Text.ToString())
                        {
                            itemIndex = i;
                            break;
                        }
                    }
                    catch (Exception exp)
                    {
                        continue;
                    }
                }
                price.Text = Main_Form.g_Form[itemIndex][3].ToString();
            }
        }

        private void button6_Click(object sender, EventArgs e)
        {
            itemgroup.DisplayMember = "value";
            itemgroup.ValueMember = "price";
            moneyitem mi = new moneyitem();
            mi.color = ComboColor2.Text;
            mi.ability = ComboAbility2.Text;
            mi.equip = ComboEquip2.Text;
            mi.kind = ComboKind2.Text;
            mi.per = ComboPer2.Text;
            if(mi.kind!="破壞")
            mi.value = mi.color + "/" + mi.kind  + "/" + mi.equip + "/" + mi.ability;
            else
                mi.value = mi.color + "/" + mi.kind + "/" + mi.equip + "/" + mi.ability + "/" + mi.per;
            itemgroup.Items.Add(mi);
        }

        private void button5_Click(object sender, EventArgs e)
        {

        }

        private void button4_Click(object sender, EventArgs e)
        {
            if (Convert.ToInt32(comboAmount.Text) % 100 != 0 && comboReq.Text=="預支")
            {
                status.Text = "預支應以100元為一單位！";
                status.ForeColor = System.Drawing.Color.Red;
            }
            else if (Convert.ToInt32(comboAmount.Text) == 0 && comboReq.Text == "預支")
            {
                status.Text = "預支金額不能為0！";
                status.ForeColor = System.Drawing.Color.Red;

            }
            else
            {
                
                UserCredential credential;
                using (var stream =
                    new FileStream("client_secret.json", FileMode.Open, FileAccess.Read))
                {
                    credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                        GoogleClientSecrets.Load(stream).Secrets,
                        Scopes,
                        "user",
                        CancellationToken.None,
                        new FileDataStore(credPath, true)).Result;
                    Console.WriteLine("Credential file saved to: " + credPath);
                }
                var service = new SheetsService(new BaseClientService.Initializer()
                {
                    HttpClientInitializer = credential,
                    ApplicationName = ApplicationName,
                });
                SpreadsheetsResource.ValuesResource.GetRequest request =
                service.Spreadsheets.Values.Get(ServiceTotalForm, "雜物出售申請!A1:A1000");
                ValueRange response = request.Execute();
                Global Applyment = new Global();
                Applyment.g_Form = response.Values;
                int rangex = Applyment.g_Form.Count + 1;
                var oblist = new List<object>() { DateTime.Now.ToString(), textBox1.Text, comboApplier.Text, comboReq.Text, Convert.ToInt32(comboAmount.Text), comboEx.Text, Convert.ToInt32(textBox3.Text), Convert.ToInt32(exchangeamount.Text) };
                if (comboReq.Text == "預支")
                {
                    oblist[5] = "";
                    oblist[6] = "";
                }
                else
                {
                    oblist[4] = "";

                }
                ValueRange valueRange = new ValueRange();
                valueRange.MajorDimension = "ROWS";
                valueRange.Values = new List<IList<object>> { oblist };
                Console.WriteLine("value range num is now: " + valueRange.Values[0]);
                string range = "雜物出售申請!A" + rangex.ToString() + ":H" + rangex.ToString();
                SpreadsheetsResource.ValuesResource.UpdateRequest update = service.Spreadsheets.Values.Update(valueRange, ServiceTotalForm, range);
                update.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.RAW;
                UpdateValuesResponse result2 = update.Execute();
                status.Text = "申請成功!";
                status.ForeColor = System.Drawing.Color.Green;

            }
        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {
            calexchange();
        }

        private void comboEx_SelectedIndexChanged(object sender, EventArgs e)
        {
            calexchange();

        }

        private void comboAmount_TextChanged(object sender, EventArgs e)
        {

            calexchange();
        }

        private void comboReq_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboReq.SelectedIndex == 0)
            {
                comboAmount.Enabled = true;
                comboEx.Enabled = false;
                textBox3.Enabled = false;

            }
            else
            {
                comboAmount.Enabled = false;
                comboEx.Enabled = true;
                textBox3.Enabled = true;

            }
            calexchange();
        }

        private void comboAmount_Leave(object sender, EventArgs e)
        {
            
        }

        private void comboAmount_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) &&
            (e.KeyChar != '.'))
            {
                e.Handled = true;
            }
        }

        private void textBox3_KeyPress(object sender, KeyPressEventArgs e)
        {

            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) &&
            (e.KeyChar != '.'))
            {
                e.Handled = true;
            }
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {

        }
    }
}
