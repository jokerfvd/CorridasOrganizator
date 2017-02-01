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
using System.Web.Script.Serialization;
using System.Text.RegularExpressions;
using System.Threading;

namespace WindowsFormsApplication1
{
    public partial class Form1 : Form
    {
        //private WebBrowser webBrowser;
        private Dictionary<String, Corrida> corridas = new Dictionary<String, Corrida>();
        bool jaFoiMySports = false;
        bool jaFoiCorridasBr = false;

        bool fimAtivo = false;//para os que tem repetido esperar
        bool fimYes = false;
        bool fimMySports = false;

        StreamWriter fileLog;

        public Form1()
        {
            InitializeComponent();
            fileLog = new StreamWriter("log.txt", false, Encoding.UTF8);
            fileLog.WriteLine("Iniciando arquivo");
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
            if ((webBrowser.ReadyState == WebBrowserReadyState.Complete) && (webBrowser.Document.Url == e.Url))
            {
                //Application.DoEvents();

                var doc = new HtmlAgilityPack.HtmlDocument();
                //int id = int.Parse(webBrowser.Url.ToString().Split('/').Last());  
                doc.LoadHtml(webBrowser.Document.GetElementsByTagName("html")[0].OuterHtml);
                String nomeCorrida = doc.DocumentNode.SelectSingleNode("/html/body/header/div[2]/div/div/text()").InnerText.Trim().Replace("Evento selecionado:","");
                richTextBox1.Text = richTextBox1.Text + "Ativo - " + nomeCorrida + "\n";
                fileLog.WriteLine(String.Format("SiteAtivoCorrida_DocumentCompleted --> {0}", nomeCorrida));
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
                            Descricao descCombo = new Descricao("COMBO", "?????", "");
                            descCombo.modalidades.Add(new Modalidade("VEJA NO SITE", "", ""));
                            corridas[nomeCorrida].descricoes.Add(descCombo);
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

                        //pegando o preco do lote atual
                        String precoDescAte = "";
                        IEnumerable<HtmlNode> lotes = li.Descendants("ul").Where(d => d.Attributes.Contains("class") && d.Attributes["class"].Value == "lotes");
                        if (lotes != null)
                        {
                            foreach (HtmlNode lote in lotes)
                            {
                                bool next = false;
                                foreach (HtmlNode li2 in lote.SelectNodes("li"))
                                {
                                    if (li2.InnerText == "Inscrição Comum") //o próxima vai ter o valor do primeiro lote
                                        next = true;
                                    else if (next)
                                    {
                                        precoDescAte = li2.InnerText.Trim();
                                        if (precoDescAte.Contains("Até"))
                                        {
                                            //removendo o Até e multiplos espaços
                                            String[] aux2 = Regex.Replace(precoDescAte.Replace("Até ", ""), @"\s+", " ").Split(' ');
                                            if ((aux2.Length == 1) || (aux2[1] == preco))
                                                precoDescAte = aux2[0];
                                            else
                                                continue;//o valor do lote atual esta no proximo
                                        }
                                        else if (precoDescAte.Contains("Lote"))
                                        {
                                            String precoDoLote = li2.SelectSingleNode("span").InnerText;
                                            if (precoDoLote == preco)//encontrado o lote atual
                                                precoDescAte = "Fim do " + precoDescAte.Split(';')[1].Split(':')[0];
                                            else
                                                continue;
                                        }
                                        break;
                                    }
                                }
                                if (next)
                                    break;
                            }
                        }

                        Descricao descricao = new Descricao(nome, preco, precoDescAte);
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
                        corridas[nomeCorrida].descricoes.Add(descricao);
                    }
                    catch (Exception ex)
                    {
                        fileLog.WriteLine("Erro no SiteAtivoCorrida_DocumentCompleted :"+ex.Message);
                        fileLog.WriteLine("StackTrace : "+ex.StackTrace);
                    }
                }
                progressBar1.PerformStep();
                if (progressBar1.Value == progressBar1.Maximum)
                    fimAtivo = true;
            }
        }

        public void SiteAtivo_DocumentCompleted(object sender, WebBrowserDocumentCompletedEventArgs e)
        {
            WebBrowser webBrowser = (WebBrowser)sender;
            //comparo "webBrowser.Document.Url == e.Url" pois as vezes vem 2x quando tem frame
            if ((webBrowser.ReadyState == WebBrowserReadyState.Complete) && (webBrowser.Document.Url == e.Url)){
                richTextBox1.Text = richTextBox1.Text + "INICIO site ativo\n";
                //HtmlElement element = webBrowser.Document.GetElementById("modalidade_select");
                webBrowser.Document.Body.ScrollIntoView(false);

                var doc = new HtmlAgilityPack.HtmlDocument();
                doc.LoadHtml(webBrowser.Document.GetElementsByTagName("html")[0].OuterHtml);

                //selecionando somente corridas das cidades do filtro
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
                            String nomeCorrida = article.SelectSingleNode("div[1]/div[1]/header/a/div[2]/h2").InnerText.Trim();
                            String nome = nomeCorrida.Replace(String.Format(" - {0}", cidade), "");
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
                            Corrida corrida = new Corrida(nome, cidade, data, url, local, retiradaKit, encerra);
                            corridas[nomeCorrida] = corrida;
                            total++;
                            //indo para o link da corrida
                            WebBrowser wb = new WebBrowser();
                            wb.ScriptErrorsSuppressed = true;
                            wb.DocumentCompleted += new WebBrowserDocumentCompletedEventHandler(SiteAtivoCorrida_DocumentCompleted);
                            wb.Navigate("https://checkout.ativo.com/evento/" + id.ToString());
                        }
                    }
                    catch (Exception ex)
                    {
                        fileLog.WriteLine("Erro no SiteAtivo_DocumentCompleted :"+ex.Message);
                        fileLog.WriteLine("StackTrace : "+ex.StackTrace);
                    }
                }
                progressBar1.Maximum = total;
                progressBar1.Step = 1;
                richTextBox1.Text = richTextBox1.Text + "FIM site ativo - "+total.ToString()+" corridas.\n";
            }
        }

        public void SiteYes_DocumentCompleted(object sender, WebBrowserDocumentCompletedEventArgs e)
        {
            WebBrowser webBrowser = (WebBrowser)sender;
            if ((webBrowser.ReadyState == WebBrowserReadyState.Complete) && (webBrowser.Document.Url == e.Url))
            {
                richTextBox1.Text = richTextBox1.Text + "INICIO site yes\n";
                HtmlWindow frame = webBrowser.Document.Window.Frames[0];
                var doc = new HtmlAgilityPack.HtmlDocument();
                doc.LoadHtml(frame.Document.Body.OuterHtml);

                HtmlNodeCollection trs = doc.DocumentNode.SelectNodes("/body/table/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr");
                int num = 0;
                progressBar2.Maximum = trs.Count;
                progressBar2.Step = 1;

                //selecionando somente corridas dos estados do filtro
                String[] estados = textBox2.Text.Split(';');

                foreach (HtmlNode tr in trs)
                {
                    if (tr.Attributes.Contains("class") && tr.Attributes["class"].Value == "hover-link-calendar")
                    {
                        String nome = tr.SelectSingleNode("td[2]").InnerText;
                        if (!nome.ToLower().Contains("kit")){
                            String aux = tr.SelectSingleNode("td[3]").InnerText;
                            String estado = "";
                            for (int i = 0; i < estados.Length;i++)
                            {
                                if (aux.Contains(estados[i])){
                                    estado = estados[i];
                                    break;
                                }
                            }
                            if (estado != "")//se for verdade eh pq encontrou em algum da lista de estados
                            {
                                String cidade = aux.Replace("-" + estado, "");
                                DateTime data = DateTime.Parse(tr.SelectSingleNode("td[1]").InnerText);
                                String url = tr.Attributes["onclick"].Value;
                                url = url.Replace("window.open('", "");
                                url = url.Replace("','','')", "");
                                if (!url.Contains("http"))
                                    url = "";
                                Corrida corrida = new Corrida(nome, cidade, data, url, "", "", data);
                                corridas[nome] = corrida;
                                num++;
                            }
                            else
                                progressBar2.Maximum = progressBar2.Maximum - 1;//dominuindo o máximo
                        }
                    }
                    progressBar2.PerformStep();
                }
                richTextBox1.Text = richTextBox1.Text + "FIM site yes - " + num.ToString() + " corridas.\n";
                fimYes = true;
            }
        }

        public void SiteMySportsDescricao_DocumentCompleted(object sender, WebBrowserDocumentCompletedEventArgs e)
        {
            WebBrowser webBrowser = (WebBrowser)sender;
            if ((webBrowser.ReadyState == WebBrowserReadyState.Complete) && (webBrowser.Document.Url == e.Url))
            {
                var doc = new HtmlAgilityPack.HtmlDocument();
                webBrowser.Document.Body.ScrollIntoView(false);
                doc.LoadHtml(webBrowser.Document.GetElementsByTagName("html")[0].OuterHtml);
                //Application.DoEvents();
                String nomeCorrida = doc.DocumentNode.SelectSingleNode("//*[@id='divEventoDestaqueNome']/span[1]").InnerText.Trim();

                String nomeDesc = doc.DocumentNode.SelectSingleNode("//*[@id='divEventoKitInfo']/span").InnerText.Trim();
                String preco = doc.DocumentNode.Descendants("b").Last().NextSibling.InnerText.Trim();
                Descricao descricao = new Descricao(nomeDesc, preco, "");
                descricao.modalidades.Add(new Modalidade("----------", "", ""));
                corridas[nomeCorrida].descricoes.Add(descricao);
            }
        }

        public void SiteMySportsCorrida_DocumentCompleted(object sender, WebBrowserDocumentCompletedEventArgs e)
        {
            WebBrowser webBrowser = (WebBrowser)sender;
            if ((webBrowser.ReadyState == WebBrowserReadyState.Complete) && (webBrowser.Document.Url == e.Url))
            {
                var doc = new HtmlAgilityPack.HtmlDocument();
                webBrowser.Document.Body.ScrollIntoView(false);
                doc.LoadHtml(webBrowser.Document.GetElementsByTagName("html")[0].OuterHtml);
                //Application.DoEvents();
                String nomeCorrida = doc.DocumentNode.SelectSingleNode("//*[@id='divEventoDestaqueNome']/span[1]").InnerText.Trim();
                richTextBox1.Text = richTextBox1.Text + "MySports - " + nomeCorrida + "\n";

                HtmlNodeCollection kits = doc.DocumentNode.SelectNodes("//*[@id='navEvento']/ul/li[3]/ul/li");
                if (kits != null)
                {
                    foreach (HtmlNode kit in kits)
                    {
                        String url2 = kit.SelectSingleNode("a").Attributes["href"].Value;
                        WebBrowser wb = new WebBrowser();
                        wb.ScriptErrorsSuppressed = true;
                        wb.DocumentCompleted += new WebBrowserDocumentCompletedEventHandler(SiteMySportsDescricao_DocumentCompleted);
                        wb.Navigate("https://www.mysports.com.br/" + url2);
                    }
                }
                else //unico
                {
                    String url2 = doc.DocumentNode.SelectSingleNode("//*[@id='navEvento']/ul/li[3]/a").Attributes["href"].Value;
                    WebBrowser wb = new WebBrowser();
                    wb.ScriptErrorsSuppressed = true;
                    wb.DocumentCompleted += new WebBrowserDocumentCompletedEventHandler(SiteMySportsDescricao_DocumentCompleted);
                    wb.Navigate("https://www.mysports.com.br/" + url2);
                }
                progressBar3.PerformStep();
            }
        }

        public void SiteMySports_DocumentCompleted(object sender, WebBrowserDocumentCompletedEventArgs e)
        {
            WebBrowser webBrowser = (WebBrowser)sender;
            if ((webBrowser.ReadyState == WebBrowserReadyState.Complete) && (webBrowser.Document.Url == e.Url))
            {
                if (!jaFoiMySports)
                {
                    webBrowser.Document.GetElementById("formBase").InvokeMember("submit");
                    jaFoiMySports = true;
                    return;
                }

                richTextBox1.Text = richTextBox1.Text + "INICIO site mysports\n";
                var doc = new HtmlAgilityPack.HtmlDocument();
                webBrowser.Document.Body.ScrollIntoView(false);
                doc.LoadHtml(webBrowser.Document.GetElementsByTagName("html")[0].OuterHtml);

                //selecionando somente corridas dos estados do filtro
                String[] estados = textBox3.Text.Split(';');
                HtmlNodeCollection eventos = doc.DocumentNode.SelectNodes("//*[@id='slide01']");
                int num = 0;
                progressBar3.Maximum = eventos.Count;
                progressBar3.Step = 1;
                foreach (HtmlNode evento in eventos) 
                {
                    String nome = evento.SelectSingleNode("div[2]/span[2]").InnerText;
                    String estado = "";
                    String[] aux = evento.SelectSingleNode("div[2]/span[3]").InnerText.Split('-');
                    for (int i = 0; i < estados.Length;i++)
                    {
                        if (( (aux.Length == 2) && (aux[1].Trim().Contains(estados[i])) || nome.Contains(estados[i]))){
                            estado = estados[i];
                            break;
                        }
                    }
                    if (estado != "")//se for verdade eh pq encontrou em algum da lista de estados
                    {
                        String cidade = aux[0].Trim();
                        aux = evento.SelectSingleNode("div[2]/span[1]").InnerText.Split('.');
                        DateTime data = DateTime.Parse(String.Format("{0} de {1} de {2}", aux[0], aux[1], aux[2]));
                        String url = "https://www.mysports.com.br/" + evento.SelectSingleNode("div[1]/a").Attributes["href"].Value;
                        Corrida corrida = new Corrida(nome, cidade, data, url, "", "", data);
                        corridas[nome] = corrida;
                        num++;

                        //indo para o link da corrida
                        WebBrowser wb = new WebBrowser();
                        wb.ScriptErrorsSuppressed = true;
                        wb.DocumentCompleted += new WebBrowserDocumentCompletedEventHandler(SiteMySportsCorrida_DocumentCompleted);
                        wb.Navigate(url);
                    }
                    else
                        progressBar3.Maximum = progressBar3.Maximum - 1;//diminuindo o máximo
                }
                richTextBox1.Text = richTextBox1.Text + "FIM site mysports - " + num.ToString() + " corridas.\n";
                fimMySports = true;
            }
        }

        public void SiteNit2Sports_DocumentCompleted(object sender, WebBrowserDocumentCompletedEventArgs e)
        {
            WebBrowser webBrowser = (WebBrowser)sender;
            if ((webBrowser.ReadyState == WebBrowserReadyState.Complete) && (webBrowser.Document.Url == e.Url))
            {
                richTextBox1.Text = richTextBox1.Text + "INICIO site Nit2Sports\n";
                var doc = new HtmlAgilityPack.HtmlDocument();
                webBrowser.Document.Body.ScrollIntoView(false);
                doc.LoadHtml(webBrowser.Document.GetElementsByTagName("html")[0].OuterHtml);
                int num = 0;
                HtmlNodeCollection eventos = doc.DocumentNode.SelectNodes("//*[@class='article-content']");
                progressBar4.Maximum = eventos.Count;
                progressBar4.Step = 1;
                foreach (HtmlNode evento in eventos)
                {
                    String nome = evento.SelectSingleNode("div[1]/div[1]/h3/a").InnerText;
                    String url = evento.SelectSingleNode("div[1]/div[1]/h3/a").Attributes["href"].Value;
                    String local = evento.SelectSingleNode("div[1]/div[3]/span[2]/text()").InnerText;
                    String hora = evento.SelectSingleNode("div[1]/div[3]/span[1]/span/text()").InnerText;
                    String dia = evento.SelectSingleNode("div[1]/div[3]/div/text()").InnerText;
                    String mes = evento.SelectSingleNode("div[1]/div[3]/div/span").InnerText;
                    DateTime data = DateTime.Parse(String.Format("{0}/{1}/{2} {3}", dia, mes, DateTime.Now.Year, hora));
                    Corrida corrida = new Corrida(nome, "Niterói", data, url, local, "", data);
                    corridas[nome] = corrida;
                    progressBar4.PerformStep();
                    num++;
                }
                richTextBox1.Text = richTextBox1.Text + "FIM site Nit2Sports - " + num.ToString() + " corridas.\n";
            }
        }

        public void SiteCorridasBr_DocumentCompleted(object sender, WebBrowserDocumentCompletedEventArgs e)
        {
            WebBrowser webBrowser = (WebBrowser)sender;
            if ((webBrowser.ReadyState == WebBrowserReadyState.Complete) && (webBrowser.Document.Url == e.Url))
            {
                richTextBox1.Text = richTextBox1.Text + "INICIO site corridasbr\n";
                var doc = new HtmlAgilityPack.HtmlDocument();
                webBrowser.Document.Body.ScrollIntoView(false);
                doc.LoadHtml(webBrowser.Document.GetElementsByTagName("html")[0].OuterHtml);
                int num = 0;
                HtmlNodeCollection tipos3 = doc.DocumentNode.SelectNodes("//*[@class='tipo3']");
                progressBar5.Maximum = tipos3.Count/4; // dividido por 4 pois cada 4 itens compoem a corrida
                progressBar5.Step = 1;
                String[] cidades = textBox4.Text.Split(';');

                for (int i = 0; i < tipos3.Count; i=i+4 )
                {
                    String cidade = tipos3[i+1].InnerText;
                    String nome = tipos3[i + 2].InnerText;
                    bool jaTem = false;
                    foreach(String key in corridas.Keys)
                        if (key.ToLower().Contains(nome.ToLower())){//verificando se essa corrida já foi inserida
                            jaTem = true;
                            break;
                        }
                    if (cidades.Contains(cidade) && !jaTem)
                    {
                        String url = e.Url.ToString().Replace("calendario.asp", "") + tipos3[i + 2].SelectSingleNode("a").Attributes["href"].Value;
                        DateTime data = DateTime.Parse(tipos3[i].InnerText);
                        Corrida corrida = new Corrida(nome, cidade, data, url, "", "", data);
                        String distancias = tipos3[i + 3].InnerText;
                        Descricao descricao = new Descricao(distancias, "?????", "");
                        descricao.modalidades.Add(new Modalidade("----------", "", ""));
                        corrida.descricoes.Add(descricao);
                        corridas[nome] = corrida;
                        progressBar5.PerformStep();
                        num++;
                    }
                    else
                        progressBar5.Maximum = progressBar5.Maximum - 1;
                }
                richTextBox1.Text = richTextBox1.Text + "FIM site corridasbr - " + num.ToString() + " corridas.\n";
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

        private void siteMySports()
        {
            WebBrowser webBrowser = new WebBrowser();
            webBrowser.ScriptErrorsSuppressed = true;
            webBrowser.DocumentCompleted += new WebBrowserDocumentCompletedEventHandler(SiteMySports_DocumentCompleted);
            webBrowser.Navigate("https://www.mysports.com.br/calendario");
        }

        private void siteNit2Sports()
        {
            WebBrowser webBrowser = new WebBrowser();
            webBrowser.ScriptErrorsSuppressed = true;
            webBrowser.DocumentCompleted += new WebBrowserDocumentCompletedEventHandler(SiteNit2Sports_DocumentCompleted);
            webBrowser.Navigate("http://www.nit2sports.com.br/");
        }

        private void siteCorridasBr()
        {
            WebBrowser webBrowser = new WebBrowser();
            webBrowser.ScriptErrorsSuppressed = true;
            webBrowser.DocumentCompleted += new WebBrowserDocumentCompletedEventHandler(SiteCorridasBr_DocumentCompleted);
            webBrowser.Navigate("http://www.corridasbr.com.br/rj/");
        }
        
        private void button1_Click(object sender, EventArgs e)
        {
            button1.Enabled = false;
            //desabilitando alertas de segurança
            ServicePointManager.ServerCertificateValidationCallback = new RemoteCertificateValidationCallback(ValidateServerCertificate);
//*  
            siteAtivo();
            while (!fimAtivo) Application.DoEvents();
            siteYes();
            while (!fimYes) Application.DoEvents();
            siteMySports();
            while (!fimMySports) Application.DoEvents();
            siteNit2Sports();
//*/
            siteCorridasBr();

            /*
             * http://www.corridasbr.com.br/rj/calendario.asp
             * http://runnerbrasil.com.br/ (não tem do RH, só MG e SP)
             * http://www.eventosmais.com/
             * 
             */
        }

        private void gerarXLS()
        {
            fileLog.WriteLine("GRAVANDO ARQUIVO .xls");
            richTextBox1.Text = richTextBox1.Text + "GRAVANDO ARQUIVO .xls\n";
            Microsoft.Office.Interop.Excel.Application oXL = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel._Workbook oWB;
            oXL.DisplayAlerts = false;
            DateTime hoje = DateTime.Now;
            FileInfo fi = new FileInfo(String.Format("corridas-{0}-gerado-em-{1}-{2}.xls",hoje.Year,hoje.Day,hoje.Month));
            if (fi.Exists)
                fi.Delete();
            oWB = oXL.Workbooks.Add(Missing.Value);
            Microsoft.Office.Interop.Excel._Worksheet oSheet = oWB.ActiveSheet;
            oSheet.Name = "Corridas";

            //criando sheet invisivel
            Microsoft.Office.Interop.Excel.Worksheet invisible = (Microsoft.Office.Interop.Excel.Worksheet)oXL.Worksheets.Add();
            invisible.Name = "invisivel";
            invisible.Columns[8].EntireColumn.NumberFormatLocal = "dd/mm/aaaa hh:mm";
            invisible.Visible = Microsoft.Office.Interop.Excel.XlSheetVisibility.xlSheetHidden;

            //cabecalho do arquivo
            int row = 1, col = 1;
            Microsoft.Office.Interop.Excel.Range cell, linha = (Microsoft.Office.Interop.Excel.Range)oSheet.Rows[row];
            String aux = "DATA;NOME;CIDADE;LOCAL;TIPO;SUBTIPO;PREÇO;PREÇO ATÉ;ENCERRA;RETIRADA";
            foreach (String c in aux.Split(';'))
                linha.Columns[col++] = c;
            linha.Font.Bold = true;
            //setando colunas de datas
            oSheet.Columns[1].EntireColumn.NumberFormatLocal = "dd/mm/aaaa hh:mm";
            oSheet.Columns[8].EntireColumn.NumberFormatLocal = "dd/mm/aaaa hh:mm";
            oSheet.Columns[9].EntireColumn.NumberFormatLocal = "dd/mm/aaaa hh:mm";
            row++;
            int namesCount = 1, starNameRow = 1;
            foreach (var entry in corridas.OrderBy(i => i.Value.getDate()))
            {
                linha = (Microsoft.Office.Interop.Excel.Range)oSheet.Rows[row];
                Corrida corrida = entry.Value;
                try
                {
                    oSheet.Cells[row, 1] = corrida.getDate();
                    //nome. irei colocar a URL como link
                    if (corrida.getUrl() != "")
                        oSheet.Hyperlinks.Add(linha.Columns[2], corrida.getUrl(), Type.Missing, Type.Missing, corrida.getNome());
                    else
                        oSheet.Cells[row, 2] = corrida.getNome();
                    oSheet.Cells[row, 3] = corrida.getCidade();
                    oSheet.Cells[row, 4] = corrida.getLocal();
                    if (corrida.descricoes.Count > 1 || (corrida.descricoes.Count == 1 && corrida.descricoes[0].modalidades.Count > 1))
                    {
                        String tipos = "", aux2 = "";
                        foreach (Descricao descricao in corrida.descricoes)
                        {
                            tipos = tipos + "," + descricao.getNome();
/*
                            if (corrida.descricoes.Count == 1 && descricao.modalidades.Count == 1)
                            {
                                if (descricao.getPreco() != "")
                                    oSheet.Cells[row, 7] = descricao.getPreco();
                                if (descricao.getPrecoAte() != "")
                                    oSheet.Cells[row, 8] = descricao.getPrecoAte();
                                oSheet.Cells[row, 5] = descricao.getNome();
                            }
*/ 
                            //else
                            //{
                                starNameRow = namesCount;
                                for (int i = 0; i < descricao.modalidades.Count; i++)
                                {
                                    Modalidade modalidade = descricao.modalidades.ElementAt(i);
                                    if (descricao.getPreco() != "")
                                        invisible.Cells[namesCount, 7] = descricao.getPreco();
                                    if (modalidade.getPreco() != "")
                                        invisible.Cells[namesCount, 7] = modalidade.getPreco();

                                    if (descricao.getPrecoAte() != "")
                                        invisible.Cells[namesCount, 8] = descricao.getPrecoAte();
                                    else if (modalidade.getPrecoAte() != "")
                                        invisible.Cells[namesCount, 8] = modalidade.getPrecoAte();
                                    else
                                        invisible.Cells[namesCount, 8] = "?????";
                                    invisible.Cells[namesCount++, 1] = modalidade.getNome();
                                }
                                //os names ficam sempre na 1ª coluna do invisilve sheet. Em cada linha correspondente vai conter valores associados do subtipo                       
                                aux2 = String.Format("_{0}.{1}", row, descricao.getNome()).Replace(" ", "").Replace("-", "");
                                Microsoft.Office.Interop.Excel.Name name = oSheet.Names.Add(aux2, invisible.get_Range((Microsoft.Office.Interop.Excel.Range)invisible.Cells[starNameRow, 1], (Microsoft.Office.Interop.Excel.Range)invisible.Cells[namesCount - 1, 1]));
                            //}
                        }
                        tipos = tipos.Substring(1);

                        int auxCol = 5; //para caso de erro na cols 5,6,7,8 (que costumam dar mais erros)
                        String auxEng ="", auxPt="";
                        try
                        {
                            //colocando listbox. COLUNA TIPO. auxCol = 5
                            oSheet.Cells[row, auxCol].Validation.Add(Microsoft.Office.Interop.Excel.XlDVType.xlValidateList, Microsoft.Office.Interop.Excel.XlDVAlertStyle.xlValidAlertStop,
                                Microsoft.Office.Interop.Excel.XlFormatConditionOperator.xlBetween, tipos, Type.Missing);
                            oSheet.Cells[row, auxCol].Validation.InCellDropdown = true;
                            oSheet.Cells[row, auxCol].Validation.IgnoreBlank = true;
                            oSheet.Cells[row, auxCol] = corrida.descricoes[0].getNome();
                            oSheet.Cells[row, auxCol].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightBlue);

                            //colocando listbox. COLUNA SUBTIPO.
                            auxPt  = String.Format("INDIRETO(\"_{0}.\"&SUBSTITUIR(SUBSTITUIR(E{1};\"-\";\"\");\" \";\"\"))", row.ToString(), row.ToString());

                            auxCol++;//auxCol = 6
                            oSheet.Cells[row, auxCol].Validation.Add(Microsoft.Office.Interop.Excel.XlDVType.xlValidateList, Type.Missing, Type.Missing, "=" + auxPt, Type.Missing);
                            oSheet.Cells[row, auxCol].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGray);
                            oSheet.Cells[row, auxCol] = corrida.descricoes[0].modalidades[0].getNome();

                            auxCol++;//auxCol = 7
                            oSheet.Cells[row, auxCol].FormulaLocal = "=INDIRETO(\"invisivel!G\"&(LIN(" + auxPt + ") + CORRESP(F" + row.ToString() + ";" + auxPt + ";0)-1))";
                            auxCol++;//auxCol = 8
                            oSheet.Cells[row, auxCol].FormulaLocal = "=INDIRETO(\"invisivel!H\"&(LIN(" + auxPt + ") + CORRESP(F" + row.ToString() + ";" + auxPt + ";0)-1))";  
                        }
                        catch (Exception ex)
                        {
                            System.Windows.Forms.MessageBox.Show(ex.Message);
                            fileLog.WriteLine(String.Format("Erro nas formulas da corrida {0} - coluna {1} - linha {2}. {3}", corrida.getNome(), auxCol, row, ex.Message));
                            fileLog.WriteLine("auxPt: " + auxPt);
                            fileLog.WriteLine("StackTrace : "+ex.StackTrace);
                        }
                    }
                    else
                    {//so tem 1 descricao e modalidade
                        oSheet.Cells[row, 5] = corrida.descricoes[0].getNome();
                        if (corrida.descricoes[0].getPreco() != "")
                            oSheet.Cells[row, 7] = corrida.descricoes[0].getPreco();
                        if (corrida.descricoes[0].getPrecoAte() != "")
                            oSheet.Cells[row, 8] = corrida.descricoes[0].getPrecoAte();
                    }
                    if (corrida.getDate() == corrida.getEncerraDate()) //coloquei igual pois nao encontrei
                        oSheet.Cells[row, 9] = "?????";
                    else
                        oSheet.Cells[row, 9] = corrida.getEncerraDate();
                    oSheet.Cells[row, 10] = corrida.getRetiradaKit();
                }
                catch (Exception ex)
                {
                    fileLog.WriteLine(String.Format("Erro na escrita da corrida {0} - linha {1} : {2}", corrida.getNome(), row, ex.Message));
                    fileLog.WriteLine("StackTrace : "+ex.StackTrace);
                }
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

            oWB.SaveAs(fi.FullName, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookNormal, Missing.Value, Missing.Value, Missing.Value, Missing.Value,
                Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive);
            oWB.Close();
            oXL.Quit();
            fileLog.WriteLine("FIM GRAVACAO DO ARQUIVO .xls");
            richTextBox1.Text = richTextBox1.Text + "FIM GRAVACAO DO ARQUIVO .xls";

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

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            fileLog.WriteLine("Fechando arquivo");
            fileLog.Close();
        }

        private void SalvarEmArquivo()
        {
            WriteToBinaryFile("corridas", corridas);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            button2.Enabled = false;
            SalvarEmArquivo();
            button2.Enabled = true;
        }

        private void CarregarDoArquivo()
        {
            button3.Enabled = false;
            corridas = ReadFromBinaryFile<Dictionary<String, Corrida>>("corridas");
            button3.Enabled = true;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            button1.Enabled = false;
            CarregarDoArquivo();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            button4.Enabled = false;
            gerarXLS();
            button4.Enabled = true;
        }


        /// <summary>
        /// Writes the given object instance to a binary file.
        /// <para>Object type (and all child types) must be decorated with the [Serializable] attribute.</para>
        /// <para>To prevent a variable from being serialized, decorate it with the [NonSerialized] attribute; cannot be applied to properties.</para>
        /// </summary>
        /// <typeparam name="T">The type of object being written to the XML file.</typeparam>
        /// <param name="filePath">The file path to write the object instance to.</param>
        /// <param name="objectToWrite">The object instance to write to the XML file.</param>
        /// <param name="append">If false the file will be overwritten if it already exists. If true the contents will be appended to the file.</param>
        public static void WriteToBinaryFile<T>(string filePath, T objectToWrite, bool append = false)
        {
            using (Stream stream = File.Open(filePath, append ? FileMode.Append : FileMode.Create))
            {
                var binaryFormatter = new System.Runtime.Serialization.Formatters.Binary.BinaryFormatter();
                binaryFormatter.Serialize(stream, objectToWrite);
            }
        }

        /// <summary>
        /// Reads an object instance from a binary file.
        /// </summary>
        /// <typeparam name="T">The type of object to read from the XML.</typeparam>
        /// <param name="filePath">The file path to read the object instance from.</param>
        /// <returns>Returns a new instance of the object read from the binary file.</returns>
        public static T ReadFromBinaryFile<T>(string filePath)
        {
            using (Stream stream = File.Open(filePath, FileMode.Open))
            {
                var binaryFormatter = new System.Runtime.Serialization.Formatters.Binary.BinaryFormatter();
                return (T)binaryFormatter.Deserialize(stream);
            }
        }



    }
}
