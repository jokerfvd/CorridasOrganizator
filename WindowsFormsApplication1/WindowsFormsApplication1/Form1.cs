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
                        String nome = spans[0].InnerText.Trim();
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
                            descricao.modalidades.Add(new Modalidade(mod, preco2, precoAte));
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
            siteYes();
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

            //cabecalho do arquivo
            int row = 1, col = 1;
            Microsoft.Office.Interop.Excel.Range cell, linha = (Microsoft.Office.Interop.Excel.Range)oSheet.Rows[row];
            String aux = "DATA;NOME;CIDADE;LOCAL;TIPO;SUBTIPO;PREÇO;PREÇO ATÉ;ENCERRA;RETIRADA";
            foreach (String c in aux.Split(';'))
                linha.Columns[col++] = c;
            linha.Font.Bold = true;
            row++;
            foreach (var entry in corridas.OrderBy(i => i.Value.getDate()))
            {
                linha = (Microsoft.Office.Interop.Excel.Range)oSheet.Rows[row];
                linha.Interior.Color = ColorTranslator.ToOle(Color.LightGray);
                Corrida corrida = entry.Value;
                aux = String.Format("{0};{1};{2};{3};;;;;{4};{5}", corrida.getData(), corrida.getNome(), corrida.getCidade(), corrida.getLocal(), corrida.getEncerra(), corrida.getRetiradaKit());
                col = 1;
                foreach (String c in aux.Split(';'))
                {
                    if (col == 2 && corrida.getUrl() != "")//nome. irei colocar a URL como link
                        oSheet.Hyperlinks.Add(linha.Columns[col++], corrida.getUrl(), Type.Missing, Type.Missing, c);
                    else if (col == 5 && corrida.descricoes.Count > 0)//coluna TIPO
                    {
                        String tipos = "";
                        foreach(Descricao descricao in corrida.descricoes)
                            tipos = tipos + "," + descricao.getNome();
                        tipos = tipos.Substring(1);
                        linha.Columns[col].Validation.Add(Microsoft.Office.Interop.Excel.XlDVType.xlValidateList, Microsoft.Office.Interop.Excel.XlDVAlertStyle.xlValidAlertStop,
                            Microsoft.Office.Interop.Excel.XlFormatConditionOperator.xlBetween,tipos);
                        linha.Columns[col].Validation.InCellDropdown = true;
                        oSheet.Cells[row, col] = tipos.Split(',')[0];
                        linha.Columns[col++].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(255, 217, 217, 0));
                    }
                    else
                        oSheet.Cells[row, col++] = c;
                }
                row++;
                foreach(Descricao descricao in corrida.descricoes){
                    aux = String.Format(";;;;{0};;{1}", descricao.getNome(), descricao.getPreco());
                    col = 1;
                    foreach (String c in aux.Split(';'))
                    {
                        cell = (Microsoft.Office.Interop.Excel.Range)oSheet.Cells[row, col];
                        if (col == 6)//nome
                            cell.Font.Bold = true;
                        oSheet.Cells[row, col++] = c;
                    }
                    row++;
                    foreach(Modalidade modalidade in descricao.modalidades){
                        aux = String.Format(";;;;;{0};{1};{2}", modalidade.getNome(), modalidade.getPreco(), modalidade.getPrecoAte());
                        col = 1;
                        foreach (String c in aux.Split(';'))
                        {
                            cell = (Microsoft.Office.Interop.Excel.Range)oSheet.Cells[row, col];
                            if (col == 7)//nome
                                cell.Font.Bold = true;
                            oSheet.Cells[row, col++] = c;
                        }
                        row++;
                    }
                }
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
