using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace WindowsFormsApplication1
{
    [Serializable]
    class Modalidade
    {
        private String nome;
        private String preco;
        private String precoAte;

        public Modalidade(String nome, String preco, String precoAte)
        {
            this.nome = nome;
            this.preco = preco;
            this.precoAte = precoAte;
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
