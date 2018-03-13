﻿using System;
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
        String VersionCode = "v1.4.3";
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
            string firstrun = "Firstrun.ini";
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
            if (!File.Exists(firstrun))
            {
                string readText = File.ReadAllText("update_log.txt");
                MessageBox.Show(readText,"更新公告");
               File.WriteAllText(firstrun, "");
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
            public int search_index(string search, int where,int wherefrom)
            {
                int index = 0;
                for (int i = wherefrom; i < g_Form.Count; i++)
                {
                    try
                    {
                        if (g_Form[i][where].ToString() == search)
                        {
                       
                            index = i;
                            return index;
                        
                        }
                    }
                    catch
                    {
                        continue;
                    }
                }
                return index;
            }
            public bool search_string(string search, int where, int wherefrom)
            {
                for (int i = 0; i < g_Form.Count; i++)
                {
                    try { 
                        if (g_Form[i][where].ToString() == search)
                        {
                        return true;
                        }
                    }
                    catch
                    {
                    continue;
                    }
            }
                return false;
            }
            public int[] search_nonskill(ComboBox cc,ComboBox ck,ComboBox ce)
            {
                var index= new int[2];
                var c = new string[3] { "赤紅", "青藍", "靛紫" };
                var k = new string[3] { "變化", "均衡", "熟練" };
                var e = new string[4] { "上衣", "下衣", "手套","鞋子" };
                int ces=ce.SelectedIndex, cks=ck.SelectedIndex, ccs=cc.SelectedIndex;
                index[0] = search_index(cc.Text+"/"+ck.Text, 0, index[0]);
                index[0] = search_index(ce.Text, 1, index[0]);
                index[1] = index[0];
                if (ces == ce.Items.Count-1)
                {
                    Console.WriteLine("ce chage line");
                    if (cks == ck.Items.Count - 1)
                    {
                        Console.WriteLine("ck chage line");
                        if (ccs == cc.Items.Count - 1)
                        {
                            Console.WriteLine("end chage line");
                            index[1] = g_Form.Count-1;
                            return index;
                        }
                        else
                        {
                            ccs=ccs+1;
                            cks = 0;
                        }
                    }
                    else
                    {
                        cks=cks+1;
                        ces = 0;
                    }
                }
                else
                {
                    ces=ces+1;
                }
                
                index[1] = search_index(c[ccs] + "/" + k[cks], 0, 0);
                index[1] = search_index(e[ces], 1, index[1]);
                index[1]--;
                return index;
            }
            public int[] search_skill(ComboBox cc)
            {
                var index = new int[2];
                var c = new string[3] { "赤紅", "青藍", "靛紫" };
                int ccs = cc.SelectedIndex;
                index[0] = search_index(cc.Text, 0, index[0]);
                index[1] = index[0];
                if (ccs != 2)
                {
                    ccs++;
                }
                else
                {
                    index[1] = g_Form.Count-1;
                    return index;
                }
                index[1] = search_index(c[ccs], 0, index[0]);
                index[1]--;
                return index;
            }
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
        public void Upload(List<IList<object>> ob,string range)
        {
            var service = new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });
            ValueRange valueRange = new ValueRange();
            valueRange.MajorDimension = "ROWS";
            valueRange.Values = ob;
            Console.WriteLine(valueRange.Values[0][0]);
            SpreadsheetsResource.ValuesResource.UpdateRequest update = service.Spreadsheets.Values.Update(valueRange, ID_Var, range);
            update.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.RAW;
            UpdateValuesResponse result2 = update.Execute();
        }
        public void GenerateItemList(Global Main_Form, Global Skill_Form, ComboBox cc, ComboBox ce, ComboBox ck, ComboBox ca)
        {
            var search = new int[2];
            if (ck.Text != "破壞")
            {
                search = Main_Form.search_nonskill(cc, ck, ce);
                    ca.Items.Clear();

                    for (int i = search[0]; i <= search[1]; i++)
                        try
                        {
                            ca.Items.Add(Main_Form.g_Form[i][2]);
                        }
                        catch (Exception e)
                        {
                            continue;
                        }
                    if (ca.Items.Count > 0)
                        ca.SelectedIndex = 0;
            }

            else
            {
                search = Skill_Form.search_skill(cc);
                ca.Items.Clear();
                for (int i = search[0]; i <= search[1]; i++)
                    try
                    {
                        if (Skill_Form.g_Form[i][1].ToString() != "")
                        {
                            ca.Items.Add(Skill_Form.g_Form[i][1]);
                        }
                    }
                    catch (Exception e)
                    {
                        continue;
                    }
                if (ca.Items.Count > 0)
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
            bool is_checked = false;
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
                                Main_Form.g_Form[i][3] = Service_Form.g_Form[i][4];
                                Console.WriteLine("偵測到價格異動 patching... x:" + i);
                            is_checked = true;

                            }
                    }
                    catch(Exception er)
                    {
                        MessageBox.Show(er.ToString());
                        continue;
                    }
            }

            if (is_checked)
            {
                ValueRange valueRange = new ValueRange();
                valueRange.Values = new List<IList<object>> { };
                for (int i = 2; i < Main_Form.g_Form.Count; i++)
                {
                    try
                    {
                        var oblist = new List<object>() { Main_Form.g_Form[i][3] };
                        valueRange.Values.Add(oblist);
                    }
                    catch
                    {
                        var oblist = new List<object>() { };
                        valueRange.Values.Add(oblist);
                    }
                }
                string range = "非技能類!D3:D376";
                valueRange.MajorDimension = "ROWS";
                Console.WriteLine("Uploading Array...");
                SpreadsheetsResource.ValuesResource.UpdateRequest update = service.Spreadsheets.Values.Update(valueRange, ID_Var, range);
                update.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.RAW;
                UpdateValuesResponse result2 = update.Execute();
            }

            
                    if (Service_Form.search_string("後端版本號",0,375))
                    {
                         OnlineVersionCode = Service_Form.g_Form[Service_Form.search_index("後端版本號", 0,375)][1].ToString();
                    }
            


            if (VersionCode.StartsWith(OnlineVersionCode))
            {
                var oblist = new List<object>() { };
                for (int i = 5; i < Main_Form.g_Form[1].Count; i++)
                {
                    if (Main_Form.g_Form[1][i].ToString() != "")
                    {
                        oblist.Add(Main_Form.g_Form[1][i]);
                        oblist.Add("");
                        oblist.Add("");
                        oblist.Add("");
                        ComboPlayer.Items.Add(Main_Form.g_Form[1][i]);
                        ComboPlayer2.Items.Add(Main_Form.g_Form[1][i]);
                        comboApplier.Items.Add(Main_Form.g_Form[1][i]);
                    }
                    else
                        break;
                }
                if(oblist!= null)
                {
                    Upload(new List<IList<object>> { oblist }, "技能類!I2:AF2");
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
            if(ComboAbility.Text.ToString()== "指令攻擊傷害增加")
            {
                ComboPer.Items.Add("0+");
            }
            var search = new int[2];
            if (ComboKind.Text.ToString() == "破壞")
            {

                for (int i = 0; i < Skill_Form.g_Form.Count; i++)
                {
                    try
                    {
                        if (Skill_Form.g_Form[i][1].ToString() == ComboAbility.Text.ToString())
                        {
                            search[0] = i;
                            break;
                        }
                    }
                    catch(Exception exr)
                    {
                        continue;
                    }
                }
                price.Text = Skill_Form.g_Form[search[0]][2].ToString();
            }
            else
            {
                for (int i = 0; i < Main_Form.g_Form.Count; i++)
                {
                    try
                    {
                        if (Main_Form.g_Form[i][2].ToString() == ComboAbility.Text.ToString())
                        {
                            search[0] = i;
                            break;
                        }
                    }
                    catch(Exception exp)
                    {
                        continue;
                    }
                }
                price.Text = Main_Form.g_Form[search[0]][3].ToString();
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
                service.Spreadsheets.Values.Get(ID_Var,"非技能類!A1:AZ2000");
                ValueRange response = request.Execute();
                Main_Form.g_Form = response.Values;
                request =
                service.Spreadsheets.Values.Get(ID_Var, "技能類!A1:AZ2000");
                response = request.Execute();
                Skill_Form.g_Form = response.Values;
                request =
                service.Spreadsheets.Values.Get(ServiceTotalForm, "非技能類!A1:AZ2000");
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

                var search = new int[2];

                if (ComboKind.Text.ToString() == "破壞")
                {
                    search[0] = Skill_Form.search_index(ComboColor.Text, 0, 0);
                    search[0] = Skill_Form.search_index(ComboAbility.Text,1,search[0]);
                    search[0] = Skill_Form.search_index(ComboPer.Text, 2, search[0]);
                    
                    for (int i = 8; i < Skill_Form.g_Form[1].Count; i++)
                    {
                        if (Skill_Form.g_Form[1][i].ToString() == ComboPlayer.Text.ToString())
                        {
                            search[1] = i;
                            break;
                        }
                    }
                    for (int i = search[1]; i < Skill_Form.g_Form[2].Count; i++)
                    {
                        if (Skill_Form.g_Form[2][i].ToString() == ComboEquip.Text.ToString())
                        {
                            search[1] = i;
                            break;
                        }
                    }
                    prev_index = search[0];
                    prev_indexPlayer = search[1];
                    ///////
                    string parsepost = "";
                    search[0]++;
                    search[1]++;
                    
                    if (search[1] > 26)
                    {
                        parsepost += "A";
                        search[1] -= 26;
                    }

                    parsepost += Convert.ToChar(search[1] + 'A' - 1);


                    var service = new SheetsService(new BaseClientService.Initializer()
                    {
                        HttpClientInitializer = credential,
                        ApplicationName = ApplicationName,
                    });

                    string range = "技能類!" + parsepost + search[0].ToString();
                    Console.WriteLine(range);
                    int stock_num = Convert.ToInt32(Skill_Form.g_Form[search[0] - 1][search[1] - 1]);
                    stock_num++;
                    Console.WriteLine("updated stock is now: " + stock_num.ToString());
                    Skill_Form.g_Form[search[0] - 1][search[1] - 1] = stock_num.ToString();
                    Console.WriteLine("updated skill form is now: " + Skill_Form.g_Form[search[0] - 1][search[1] - 1]);
                    position.Text = range;
                    num_temp.Text = stock_num.ToString();
                    var oblist = new List<object>() { stock_num };
                    Upload(new List<IList<object>> { oblist }, range);

                }
                else
                {
                    for (int i = 0; i < Main_Form.g_Form.Count; i++)
                    {
                        try
                        {
                            if (Main_Form.g_Form[i][0].ToString() == ComboColor.Text.ToString() + "/" + ComboKind.Text.ToString())
                            {
                                search[0] = i;
                                break;
                            }
                        }
                        catch
                        {
                            continue;
                        }
                    }
                    for (int i = search[0]; i < Main_Form.g_Form.Count; i++)
                    {
                        try
                        {
                            if (Main_Form.g_Form[i][1].ToString() == ComboEquip.Text.ToString())
                            {
                                search[0] = i;
                                break;
                            }
                        }
                        catch
                        {
                            continue;
                        }
                    }
                    for (int i = search[0]; i < Main_Form.g_Form.Count; i++)
                        {
                            try
                            {
                                if (Main_Form.g_Form[i][2].ToString() == ComboAbility.Text.ToString())
                                {
                                    search[0] = i;
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
                            search[1] = i;
                            break;
                        }
                    }
                    prev_index = search[0];
                    prev_indexPlayer = search[1];


                    //////
                    string parsepost = "";
                    search[1]++;
                    search[0]++;


                    if (search[1] > 26)
                    {
                        parsepost += "A";
                        search[1] -= 26;
                    }
                    
                    var service = new SheetsService(new BaseClientService.Initializer()
                    {
                        HttpClientInitializer = credential,
                        ApplicationName = ApplicationName,
                    });

                    parsepost += Convert.ToChar(search[1]+'A'-1);
                    string range = "非技能類!"+parsepost+search[0].ToString();
                    Console.WriteLine(range);

                    int stock_num = Convert.ToInt32(Main_Form.g_Form[search[0] - 1][search[1] - 1]);
                    stock_num++;
                    Main_Form.g_Form[search[0] - 1][search[1]-1] = stock_num.ToString();
                    position.Text = range;
                    num_temp.Text = stock_num.ToString();
                    var oblist = new List<object>() { stock_num };
                    Upload(new List<IList<object>> { oblist }, range);

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
                Upload(new List<IList<object>> { oblist }, position.Text);
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
            var search = new int[2];
            if (ComboKind.Text.ToString() == "破壞")
            {

                for (int i = 0; i < Skill_Form.g_Form.Count; i++)
                {
                    try
                    {
                        if (Skill_Form.g_Form[i][1].ToString() == ComboAbility.Text.ToString())
                        {
                            search[0] = i;
                            break;
                        }
                    }
                    catch (Exception exr)
                    {
                        continue;
                    }
                }
                price.Text = Skill_Form.g_Form[search[0]][2].ToString();
            }
            else
            {
                for (int i = 0; i < Main_Form.g_Form.Count; i++)
                {
                    try
                    {
                        if (Main_Form.g_Form[i][2].ToString() == ComboAbility.Text.ToString())
                        {
                            search[0] = i;
                            break;
                        }
                    }
                    catch (Exception exp)
                    {
                        continue;
                    }
                }
                price.Text = Main_Form.g_Form[search[0]][3].ToString();
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
                string range = "雜物出售申請!A" + rangex.ToString() + ":H" + rangex.ToString();
                Upload(new List<IList<object>> { oblist }, range);
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

        private void ComboColor_SelectionChangeCommitted(object sender, EventArgs e)
        {

        }
    }
}
