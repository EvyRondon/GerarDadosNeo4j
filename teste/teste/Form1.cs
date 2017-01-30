using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace teste
{
    public partial class Form1 : Form
    {
        private int pessoaId { get; set; }
        private int artigoId { get; set; }
        private int palavraId { get; set; }

        List<Autor> autores;
        List<PalavraChave> palavrasChaves;
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            try
            {
                gerarCarga();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Erro ao executar Leitura do Arquivo" + ex);
                throw;
            }
        }
        public void gerarCarga()
        {
            autores = new List<Autor>();
            palavrasChaves = new List<PalavraChave>();
            //StreamReader rd = new StreamReader(@"E:\Arquivos\Imagens\Teste\2014-1_2014-5(modificado2).csv", Encoding.GetEncoding("iso-8859-1"));
            StreamReader rd = new StreamReader(@"C:\Users\29888\Pictures\TesteNeo\2014-1_2014-5(modificado2).csv", Encoding.GetEncoding("iso-8859-1"));

            string linha = null;

            pessoaId = 0;
            artigoId = 0;
            palavraId = 0;

            while ((linha = rd.ReadLine()) != null)
            {
                string[] linhaSeparada = linha.Split(';');

                List<string> listaCompletaEmail = RetirarEmailDuplicado(linhaSeparada);
                List<string> listaNome = linhaSeparada[1].Replace('[', ' ').Replace(']', ' ').Trim().Split(',').ToList();
                List<string> listaPalavraChave = linhaSeparada[2].Replace(" and ", " , ").Replace('.', ' ').Split(',').ToList();

                GerarPalavraChave(listaPalavraChave);
                GerarAutor(listaNome, listaCompletaEmail);
                CriaArtigoCSV(linhaSeparada[0].ToString(), linhaSeparada[2].ToString(), linhaSeparada[4].ToString());
            }
            rd.Close();
        }
        public void GerarPalavraChave(List<string> listaPalavraChaves)
        {
            foreach (string palavraChave in listaPalavraChaves)
            {
                string palavraChaveFormatada = RemoveAccents(palavraChave.ToLower().Trim());
                if (!palavrasChaves.Any(x => x.Descricao == palavraChaveFormatada)){
                    var palavra = new PalavraChave { ID = palavraId, Descricao = palavraChaveFormatada };
                    palavrasChaves.Add(palavra);
                    CriaPalavraChaveCSV(palavra);
                    CriarArtigoPalavraChaveCSV();
                    palavraId++;
                }else
                {
                    var idPalavraRepetida = palavrasChaves.Where(x => x.Descricao == palavraChaveFormatada).FirstOrDefault().ID;
                    CriarArtigoPalavraChaveCSV(idPalavraRepetida);
                }
            }
        }
        public void GerarAutor(List<string> nomeAutorCompleto, List<string> listaCompletaEmail)
        {
            foreach (string Email in listaCompletaEmail)
            {
                string nome = "";
                string email = Email;
                foreach (var Nome in nomeAutorCompleto)
                {
                    nome = Nome;
                    if (!autores.Any(x => x.Email == Email))
                    {
                        autores.Add(new Autor { idAutor = pessoaId, Nome = nome, Email = email });
                        CriaAutorCSV(nome, email);
                    }
                    else
                    {
                        int idPessoaCadastrada = autores.Where(x => x.Email == Email).FirstOrDefault().idAutor;
                        CriaArtigoAutorCSV("CoAutor", idPessoaCadastrada);
                    }
                    nomeAutorCompleto.Remove(nome);
                    break;
                }
            }
        }
        public List<string> RetirarEmailDuplicado(string[] linhaSeparada)
        {
            string[] linhaEmail = null;
            List<string> listaEmail = new List<string>();
            linhaEmail = linhaSeparada[3].Replace('[', ' ').Replace(']', ' ').Trim().Split(',');

            foreach (var item in linhaEmail)
            {
                var novoItem = item.Trim();
                var rbmet = "'rbmet@rbmet.org.br'";
                if (!listaEmail.Contains(novoItem))
                    listaEmail.Add(novoItem);
                else if (listaEmail.Contains(rbmet))
                    listaEmail.Remove(rbmet);
            }
            return listaEmail;
        }
        public void CriaAutorCSV(string nome, string email)
        {
            //using (TextWriter sw = new StreamWriter(@"E:\Arquivos\Imagens\Teste\Autor.csv", true, Encoding.GetEncoding("iso-8859-1")))
            using (TextWriter sw = new StreamWriter(@"C:\Users\29888\Pictures\TesteNeo\Autor.csv", true, Encoding.GetEncoding("iso-8859-1")))
            {
                string strNome = nome;
                string strEmail = email;//Note it's a float not string
                sw.WriteLine("{0};{1};{2}", pessoaId, strNome, strEmail);
                CriaArtigoAutorCSV("Escreveu");
                pessoaId++;
            }
        }
        public void CriaArtigoCSV(string titulo, string palavras, string ano)
        {
            //using (TextWriter sw = new StreamWriter(@"E:\Arquivos\Imagens\Teste\Artigo.csv", true, Encoding.GetEncoding("iso-8859-1")))
            using (TextWriter sw = new StreamWriter(@"C:\Users\29888\Pictures\TesteNeo\Artigo.csv", true, Encoding.GetEncoding("iso-8859-1")))
            {
                sw.WriteLine("{0};{1};{2};{3}", artigoId, titulo, palavras, ano);
                artigoId++;
            }
        }
        public void CriaArtigoAutorCSV( string tipo, int? idPessoaCadastrada = null)
        {
            //using (TextWriter sw = new StreamWriter(@"E:\Arquivos\Imagens\Teste\ArtigoAutor.csv", true, Encoding.GetEncoding("iso-8859-1")))
            using (TextWriter sw = new StreamWriter(@"C:\Users\29888\Pictures\TesteNeo\ArtigoAutor.csv", true, Encoding.GetEncoding("iso-8859-1")))
            {
                sw.WriteLine("{0};{1};{2}", idPessoaCadastrada ?? pessoaId, artigoId, tipo);
            }
        }
        public void CriaPalavraChaveCSV(PalavraChave palavra)
        {
            //using (TextWriter sw = new StreamWriter(@"E:\Arquivos\Imagens\Teste\PalavraChave.csv", true, Encoding.GetEncoding("iso-8859-1")))
            using (TextWriter sw = new StreamWriter(@"C:\Users\29888\Pictures\TesteNeo\PalavraChave.csv", true, Encoding.GetEncoding("iso-8859-1")))
            {
                sw.WriteLine("{0};{1}", palavra.ID, palavra.Descricao);
            }
        }
        public void CriarArtigoPalavraChaveCSV(int? idPalavraRepetida = null)
        {
            //using (TextWriter sw = new StreamWriter(@"E:\Arquivos\Imagens\Teste\ArtigoPalavraChave.csv", true, Encoding.GetEncoding("iso-8859-1")))
            using (TextWriter sw = new StreamWriter(@"C:\Users\29888\Pictures\TesteNeo\ArtigoPalavraChave.csv", true, Encoding.GetEncoding("iso-8859-1")))
            {
                sw.WriteLine("{0};{1}", artigoId, idPalavraRepetida ?? palavraId);
            }
        }

        public string RemoveAccents(string text)
        {
            StringBuilder sbReturn = new StringBuilder();
            var arrayText = text.Normalize(NormalizationForm.FormD).ToCharArray();
            foreach (char letter in arrayText)
            {
                if (CharUnicodeInfo.GetUnicodeCategory(letter) != UnicodeCategory.NonSpacingMark)
                    sbReturn.Append(letter);
            }
            return sbReturn.ToString();
        }
    }
}
