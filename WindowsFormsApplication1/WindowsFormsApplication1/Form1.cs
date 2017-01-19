using HtmlAgilityPack;
using mshtml;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Security;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace WindowsFormsApplication1
{
    public partial class Form1 : Form
    {
        //private WebBrowser webBrowser;
        private Dictionary<int, Corrida> corridas = new Dictionary<int, Corrida>();

        StreamWriter fileLog;

        public Form1()
        {
            InitializeComponent();
            fileLog = new StreamWriter("log.txt", false, Encoding.UTF8);
        }

        private String englishMonth(String data){
	        try{
                data = data.Replace("Fev","Feb");
                data = data.Replace("Abr","Apr");
                data = data.Replace("Mai","May");
                data = data.Replace("Ago","Aug");
                data = data.Replace("Set","Sep");
                data = data.Replace("Out","Oct");
                data = data.Replace("Dez","Dec");
	        }
	        catch (Exception){}
            return data;
        }

        private String getIdDaCorrida(String url){
            String[] aux = url.Split('/');
            for (int i=0; i < aux.Length;i++){
                if (aux[i].Contains("corrida-de"))
                    return aux[i+1];
            }
	        return null;
        }

        public void SiteAtivoCorrida_DocumentCompleted(object sender, WebBrowserDocumentCompletedEventArgs e)
        {
            WebBrowser webBrowser = (WebBrowser)sender;
            if (webBrowser.ReadyState == WebBrowserReadyState.Complete)
            {
                Application.DoEvents();

                var doc = new HtmlAgilityPack.HtmlDocument();
                int id = int.Parse(webBrowser.Url.ToString().Split('/').Last());
                richTextBox1.Text = richTextBox1.Text + "PROCESSANDO " + corridas[id].getNome() + "\n";
                doc.LoadHtml(webBrowser.Document.GetElementsByTagName("html")[0].OuterHtml);
                fileLog.WriteLine(String.Format("SiteAtivoCorrida_DocumentCompleted --> {0}", id));
                foreach (HtmlNode li in doc.DocumentNode.SelectNodes("/html/body/main/div/div/div[1]/ul/li")){
                    try
                    {
                        HtmlNode aux = null;
                        try
                        {
                            aux = li.Descendants("div").Where(d => d.Attributes.Contains("class") && d.Attributes["class"].Value == "descricao").First();
                        }
                        catch (Exception)//
                        {
                            fileLog.WriteLine("possível combo");
                            corridas[id].descricoes.Add(new Descricao("COMBO","?????"));
                            continue;
                        }
                        HtmlNodeCollection spans = aux.SelectNodes("span");
                        String nome = spans[0].InnerText.Trim().Replace(" ", "");
                        String preco = "";
                        if (li.Descendants("p").ToList().Count > 0)
                        {
                            aux = li.Descendants("p").Where(d => d.Attributes.Contains("class") && d.Attributes["class"].Value == "b").First();
                            preco = aux.InnerText.Trim().Split(' ')[0];
                        }
                        Descricao descricao = new Descricao(nome, preco);
                        foreach (HtmlNode radio in li.Descendants("input").Where(d => d.Attributes.Contains("name") && d.Attributes["name"].Value == "modalidade"))
                        {
                            HtmlNode parent = radio.ParentNode;
                            String mod = null;
                            if (parent.SelectNodes("label") != null)
                                mod = parent.SelectSingleNode("label").InnerText;
                            else if (parent.SelectNodes("span") != null)
                                mod = parent.SelectSingleNode("span").InnerText;
                            String preco2 = "",precoAte = "";
                            if (li.Descendants("b").ToList().Count > 0)
                            {
                                HtmlNode aux2 = li.Descendants("b").Where(d => d.Attributes.Contains("class") && d.Attributes["class"].Value == "b").First();
                                preco2 = aux2.InnerText.Trim().Split(' ')[0];
                                if (!preco2.Contains("R$"))//tem que ter R$ ou é lixo
                                    preco2 = "";
                                parent = aux2.ParentNode;
                                aux2 = parent.SelectSingleNode("span");
                                if (aux2 != null)
                                {
                                    precoAte = aux2.InnerText.Trim();
                                    precoAte = precoAte.Replace("Até ", "");
                                }
                            }
                            descricao.modalidades.Add(new Modalidade(mod.Trim(), preco2, precoAte));
                        }
                        corridas[id].descricoes.Add(descricao);
                    }
                    catch (Exception ex)
                    {
                        fileLog.WriteLine(ex.Message);
                    }
                }
                progressBar1.PerformStep();
            }
        }

        public void SiteAtivo_DocumentCompleted(object sender, WebBrowserDocumentCompletedEventArgs e)
        {
            WebBrowser webBrowser = (WebBrowser)sender;
            if (webBrowser.ReadyState == WebBrowserReadyState.Complete){
                richTextBox1.Text = richTextBox1.Text + "INICIO site ativo\n";
                HtmlElement element = webBrowser.Document.GetElementById("modalidade_select");
                webBrowser.Document.Body.ScrollIntoView(false);

                var doc = new HtmlAgilityPack.HtmlDocument();
                doc.LoadHtml(webBrowser.Document.GetElementsByTagName("html")[0].OuterHtml);

                String[] cidades = textBox1.Text.Split(';');

                int total = 0;
                HtmlNodeCollection articles = doc.DocumentNode.SelectNodes("//*[@id='container']/article");
                foreach (HtmlNode article in articles) 
                {
                    try
                    {
                        String cidade = article.SelectSingleNode("div[1]/div[1]/header/a/div[4]/div/span").InnerText;
                        if (cidades.Contains(cidade))
                        {
                            String nome = article.SelectSingleNode("div[1]/div[1]/header/a/div[2]/h2").InnerText;
                            nome = nome.Replace(String.Format(" - {0}", cidade), "");
                            HtmlNodeCollection aux = article.SelectNodes("div[1]/div[1]/header/a/time/span");
                            DateTime data = DateTime.Parse(englishMonth(String.Format("{0}/{1}/{2}", aux[1].InnerText, aux[2].InnerText, aux[0].InnerText)));
                            String url = article.SelectSingleNode("div[1]/figure/a").Attributes["href"].Value;
                            if (!url.Contains("corrida-de"))
                            {
                                fileLog.WriteLine("DESCARTADO");
                                continue;
                            }
                            int id = int.Parse(getIdDaCorrida(url));

                            //pegando alguns dados da corrida
                            HtmlAgilityPack.HtmlWeb web = new HtmlWeb();
                            HtmlAgilityPack.HtmlDocument doc2 = web.Load(url);
                            String local, largada, retiradaKit;
                            DateTime encerra;
                            local = doc2.DocumentNode.SelectSingleNode("//*[@id='main']/section/div/div/div[2]/div/div/div[1]/div[1]/div[2]/div[3]/p").InnerText;
                            local = local.Replace(String.Format("Brasil - {0} - ",cidade), "");
                            largada = doc2.DocumentNode.SelectSingleNode("//*[@id='main']/section/div/div/div[2]/div/div/div[1]/div[1]/div[3]/div[3]/p").InnerText;
                            retiradaKit = doc2.DocumentNode.SelectSingleNode("//*[@id='main']/section/div/div/div[2]/div/div/div/div[1]/div[4]/div[3]/p").InnerText.Trim();
                            retiradaKit = Regex.Replace(retiradaKit, @"\t|\n|\r", "");
                            encerra = DateTime.Parse(doc2.DocumentNode.SelectSingleNode("//*[@id='main']/section/div/div/div[2]/div/div/div/p[1]").InnerText);
                            data = DateTime.Parse(data.ToString("dd/MM/yyyy ") + largada);
                            Corrida corrida = new Corrida(id, nome, cidade, data, url, local, retiradaKit, encerra);
                            corridas[id] = corrida;

                            total++;
                            //indo para o link da corrida
                            WebBrowser wb = new WebBrowser();
                            wb.ScriptErrorsSuppressed = true;
                            wb.DocumentCompleted += new WebBrowserDocumentCompletedEventHandler(SiteAtivoCorrida_DocumentCompleted);
                            wb.Navigate("https://checkout.ativo.com/evento/" + id.ToString());

                            ///////////////////////////////////APAGAR
                            if (total > 1)
                                break;
                        }
                    }
                    catch (Exception ex)
                    {
                        fileLog.WriteLine(ex.Message);
                    }
                }
                progressBar1.Maximum = total;
                progressBar1.Step = 1;
                richTextBox1.Text = richTextBox1.Text + "FIM site ativo\n";
            }
        }

        public void SiteYes_DocumentCompleted(object sender, WebBrowserDocumentCompletedEventArgs e)
        {
            WebBrowser webBrowser = (WebBrowser)sender;
            if (webBrowser.ReadyState == WebBrowserReadyState.Complete)
            {
                richTextBox1.Text = richTextBox1.Text + "INICIO site yes\n";
                HtmlWindow frame = webBrowser.Document.Window.Frames[0];
                var doc = new HtmlAgilityPack.HtmlDocument();
                doc.LoadHtml(frame.Document.Body.OuterHtml);

                HtmlNodeCollection trs = doc.DocumentNode.SelectNodes("/body/table/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr");
                int num = 100;
                progressBar2.Maximum = trs.Count;
                progressBar2.Step = 1;
                foreach (HtmlNode tr in trs)
                {
                    if (tr.Attributes.Contains("class") && tr.Attributes["class"].Value == "hover-link-calendar")
                    {
                        String nome = tr.SelectSingleNode("td[2]").InnerText;
                        String cidade = tr.SelectSingleNode("td[3]").InnerText;
                        if (cidade.Contains("RJ") && !nome.Contains("Kit"))
                        {
                            cidade = cidade.Replace("-RJ", "");
                            DateTime data = DateTime.Parse(tr.SelectSingleNode("td[1]").InnerText);
                            String url = tr.Attributes["onclick"].Value;
                            url = url.Replace("window.open('", "");
                            url = url.Replace("','','')", "");
                            if (!url.Contains("http"))
                                url = "";      
                            Corrida corrida = new Corrida(num, nome, cidade, data, url, "", "", data);
                            corridas[num++] = corrida;
                        }
                    }
                    progressBar2.PerformStep();
                }
                richTextBox1.Text = richTextBox1.Text + "FIM site yes\n";
            }
        }

        public static bool ValidateServerCertificate(object sender, X509Certificate certificate, X509Chain chain, SslPolicyErrors sslPolicyErrors)
        {
            return true;
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            
        }

        private void siteAtivo()
        {
            WebBrowser webBrowser = new WebBrowser();
            webBrowser.ScriptErrorsSuppressed = true;
            webBrowser.DocumentCompleted += new WebBrowserDocumentCompletedEventHandler(SiteAtivo_DocumentCompleted);
            webBrowser.Navigate("https://www.ativo.com/calendario/");
        }

        private void siteYes()
        {
            WebBrowser webBrowser = new WebBrowser();
            webBrowser.ScriptErrorsSuppressed = true;
            webBrowser.DocumentCompleted += new WebBrowserDocumentCompletedEventHandler(SiteYes_DocumentCompleted);
            webBrowser.Navigate("http://www.yescom.com.br/site/calendario.html");
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //desabilitando alertas de segurança
            ServicePointManager.ServerCertificateValidationCallback = new RemoteCertificateValidationCallback(ValidateServerCertificate);
            siteAtivo();
            //siteYes();
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            richTextBox1.Text = richTextBox1.Text + "GRAVANDO ARQUIVO .xls\n";
            Microsoft.Office.Interop.Excel.Application oXL = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel._Workbook oWB;
            oXL.DisplayAlerts = false;
            FileInfo fi = new FileInfo(@"corridas.xls");
            if (fi.Exists)
                fi.Delete();
            oWB = oXL.Workbooks.Add(Missing.Value);
            Microsoft.Office.Interop.Excel._Worksheet oSheet = oWB.ActiveSheet;
            oSheet.Name = "Corridas";

            //criando sheet invisivel
            Microsoft.Office.Interop.Excel.Worksheet invisible =  (Microsoft.Office.Interop.Excel.Worksheet)oXL.Worksheets.Add();
            invisible.Name = "invisivel";
            //////////////////////DEIXA INVISIVEL DEPOIS
            //invisible.Visible = Microsoft.Office.Interop.Excel.XlSheetVisibility.xlSheetHidden;

            //cabecalho do arquivo
            int row = 1, col = 1;
            Microsoft.Office.Interop.Excel.Range cell, linha = (Microsoft.Office.Interop.Excel.Range)oSheet.Rows[row];
            String aux = "DATA;NOME;CIDADE;LOCAL;TIPO;SUBTIPO;PREÇO;PREÇO ATÉ;ENCERRA;RETIRADA";
            foreach (String c in aux.Split(';'))
                linha.Columns[col++] = c;
            linha.Font.Bold = true;
            row++;
            int namesCount = 1, starNameRow = 1;
            foreach (var entry in corridas.OrderBy(i => i.Value.getDate()))
            {
                linha = (Microsoft.Office.Interop.Excel.Range)oSheet.Rows[row];
                //linha.Interior.Color = ColorTranslator.ToOle(Color.LightGray);
                Corrida corrida = entry.Value;
                oSheet.Cells[row, 1] = corrida.getData();
                //nome. irei colocar a URL como link
                oSheet.Hyperlinks.Add(linha.Columns[2], corrida.getUrl(), Type.Missing, Type.Missing, corrida.getNome());
                oSheet.Cells[row, 3] = corrida.getCidade();
                oSheet.Cells[row, 4] = corrida.getLocal();
                if (corrida.descricoes.Count > 0)//coluna TIPO
                {
                    String tipos = "";
                    foreach(Descricao descricao in corrida.descricoes){
                        tipos = tipos + "," + descricao.getNome();
                        starNameRow = namesCount;
                        for(int i=0; i < descricao.modalidades.Count; i++){
                            Modalidade modalidade = descricao.modalidades.ElementAt(i);
                            if (descricao.getPreco() != "")
                                invisible.Cells[namesCount, 7] = descricao.getPreco(); 
                            if (modalidade.getPreco() != "")
                                invisible.Cells[namesCount, 7] = modalidade.getPreco();
                            if (modalidade.getPrecoAte() != "")
                                invisible.Cells[namesCount, 8] = modalidade.getPrecoAte();
                            invisible.Cells[namesCount++, 1] = modalidade.getNome();                           
                        }
                           
                        //os names ficam sempre na 1ª coluna do invisilve sheet. Em cada linha correspondente vai conter valores associados do subtipo
                        String auxName = String.Format("#{0}.{1}.{2}", corrida.getCidade(), corrida.getData(), descricao.getNome()).Replace(" ","").Replace("/","").Replace(":","");
                        Microsoft.Office.Interop.Excel.Name name = oSheet.Names.Add(auxName, invisible.get_Range((Microsoft.Office.Interop.Excel.Range)invisible.Cells[starNameRow, 1], (Microsoft.Office.Interop.Excel.Range)invisible.Cells[namesCount - 1, 1])); 
                    }
                    tipos = tipos.Substring(1);
                    //colocando listbox. COLUNA TIPO
                    oSheet.Cells[row, 5].Validation.Add(Microsoft.Office.Interop.Excel.XlDVType.xlValidateList, Microsoft.Office.Interop.Excel.XlDVAlertStyle.xlValidAlertStop,
                        Microsoft.Office.Interop.Excel.XlFormatConditionOperator.xlBetween,tipos, Type.Missing);
                    oSheet.Cells[row, 5].Validation.InCellDropdown = true;
                    oSheet.Cells[row, 5].Validation.IgnoreBlank = true;
                    oSheet.Cells[row, 5] = corrida.descricoes[0].getNome();
                    oSheet.Cells[row, 5].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightBlue);

                    //colocando listbox. COLUNA SUBTIPO.
                    oSheet.Cells[row, 6].Validation.Add(Microsoft.Office.Interop.Excel.XlDVType.xlValidateList, Type.Missing, Type.Missing, "=INDIRETO($E$" + row.ToString() + ")", Type.Missing); //coluna TIPO (col 5, E)
                    oSheet.Cells[row, 6].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGray);
                    oSheet.Cells[row, 6] = corrida.descricoes[0].modalidades[0].getNome();

                    
                    try
                    {
                        //Na formula tem que ta em ingles! PQP! No validation nao.
                        String auxName = String.Format("E{0}&\".")
                            //PARAIE POR AQUIUUUUUUUUUUUUUUUUUUUUUUUuu
                            //=C2&DIA(A2)&MÊS(A2)&HORA(A2)
                        oSheet.Cells[row, 7].Formula = "=INDIRECT(\"invisivel!G\"&(ROW(INDIRECT(E" + row.ToString() + ")) + MATCH(F" + row.ToString() + ",INDIRECT(E" + row.ToString() + "),0)-1))";
                        oSheet.Cells[row, 8].Formula = "=INDIRECT(\"invisivel!H\"&(ROW(INDIRECT(E" + row.ToString() + ")) + MATCH(F" + row.ToString() + ",INDIRECT(E" + row.ToString() + "),0)-1))";

                        //oSheet.Cells[row, 8] = "=INDIRETO(\"invisivel!H\"&(LIN(INDIRETO(E" + row.ToString() + ")) + CORRESP(F" + row.ToString() + ",INDIRETO(E" + row.ToString() + "),0)-1))";   

                    }
                    catch (Exception ex) { System.Windows.Forms.MessageBox.Show(ex.Message); }
                    //o INDIRECT SO FUNCIONOU EM PORTUGUES
                    //=INDIRETO("invisivel!G"&LIN(INDIRETO(E2)))
                    //INDIRETO("invisivel!G"&(LIN(INDIRETO($E$2)) + CORRESP(F2;INDIRETO(E2);0)-1))
                }
                oSheet.Cells[row, 9] = corrida.getEncerra();
                oSheet.Cells[row, 10] = corrida.getRetiradaKit();            

                row++;
            }
            oSheet.Columns[1].AutoFit();
            oSheet.Columns[2].AutoFit();
            oSheet.Columns[3].AutoFit();
            oSheet.Columns[5].AutoFit();
            oSheet.Columns[6].AutoFit();
            oSheet.Columns[7].AutoFit();
            oSheet.Columns[8].AutoFit();
            oSheet.Columns[9].AutoFit();
            oSheet.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;

            fileLog.Close();

            oWB.SaveAs(fi.FullName, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookNormal, Missing.Value, Missing.Value, Missing.Value, Missing.Value,
                Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive);
            oWB.Close();
            oXL.Quit();

            //Clean up
            //NOTE: When in release mode, this does the trick
            GC.WaitForPendingFinalizers();
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();

            Marshal.FinalReleaseComObject(oSheet);
            Marshal.FinalReleaseComObject(oWB);
            Marshal.FinalReleaseComObject(oXL);
        }


    }
}
