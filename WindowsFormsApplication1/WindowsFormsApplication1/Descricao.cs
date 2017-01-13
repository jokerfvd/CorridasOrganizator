using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace WindowsFormsApplication1
{
    class Descricao
    {
        public List<Modalidade> modalidades;
        private String nome;
        private String preco;
        public Descricao(String nome, String preco)
        {
            this.nome = nome;
            this.preco = preco;
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
    }
}
