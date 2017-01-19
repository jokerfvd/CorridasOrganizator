using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace WindowsFormsApplication1
{
    class Corrida
    {
        private int id;
        private String nome;
        private String cidade;
        private DateTime data;
        private String url;
        private String local;
        private String retiradaKit;
        private DateTime encerra;

        public List<Descricao> descricoes;

        public Corrida(int id, String nome, String cidade, DateTime data, String url, String local, String retiradaKit, DateTime encerra)
        {
            this.id = id;
            this.nome = nome;
            this.cidade = cidade;
            this.data = data;
            this.url = url;
            this.local = local;
            this.retiradaKit = retiradaKit;
            this.encerra = encerra;
            descricoes = new List<Descricao>();
        }

        public int getId()
        {
            return id;
        }

        public String getData()
        {
            return data.ToString("dd/MM/yyyy HH:mm");
        }

        public DateTime getDate()
        {
            return data;
        }

        public String getCidade()
        {
            return cidade;
        }

        public String getNome()
        {
            return nome;
        }

        public String getUrl()
        {
            return url;
        }

        public String getLocal()
        {
            return local;
        }

        public String getRetiradaKit()
        {
            return retiradaKit;
        }

        public String getEncerra()
        {
            return encerra.ToString("dd/MM/yyyy");
        }
    }
}
