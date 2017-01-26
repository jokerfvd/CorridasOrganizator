using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace WindowsFormsApplication1
{
    [Serializable]
    class Descricao
    {
        public List<Modalidade> modalidades;
        private String nome;
        private String preco;
        private String precoAte;
        public Descricao(String nome, String preco, String precoAte)
        {
            this.nome = nome;
            this.preco = preco;
            this.precoAte = precoAte;
            modalidades = new List<Modalidade>();
        }

        public String getNome()
        {
            return nome;
        }

        public String getPreco()
        {
            return preco;
        }

        public String getPrecoAte()
        {
            return precoAte;
        }
    }
}
