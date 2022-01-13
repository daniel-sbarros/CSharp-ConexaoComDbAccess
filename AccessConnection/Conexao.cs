using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AccessConnection
{
    internal class Conexao
    {
        private OleDbConnection conn;

        public Conexao()
        {
            string db_name = @$"{Directory.GetCurrentDirectory()}\conexaoAccess.accdb";
            conn = new($"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={db_name};Persist Security Info=False;");
        }

        public bool open()
        {
            try
            {
                if (conn.State != System.Data.ConnectionState.Open)
                {
                    conn.Open();
                }
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }

        public bool close()
        {
            try
            {
                if (conn.State == System.Data.ConnectionState.Open)
                {
                    conn.Close();
                }
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }

        public OleDbDataReader pesquisar(string query)
        {
            OleDbCommand cmd = new(query, conn);

            try
            {
                if (open())
                {
                    OleDbDataReader reader = cmd.ExecuteReader();
                    return reader;
                }
                else
                {
                    return null;
                }
            }
            catch (OleDbException e)
            {
                Console.WriteLine("Error: {0}", e.Errors[0].Message);
                return null;
            }
        }

        public void listar(string query)
        {
            OleDbDataReader reader = pesquisar(query);

            if (reader != null)
            {
                Console.WriteLine("Os valores retornados da tabela são:");

                while (reader.Read()) Console.WriteLine($"Nome: {reader.GetString(1)}\tEmail: {reader.GetString(2)}");

                reader.Close();
                close();
            }
            else
            {
                Console.WriteLine("A pesquisa retornou valor nulo.");
            }
        }

        public bool add(string query, params string[] values)
        {
            try
            {
                if (open() && !string.IsNullOrEmpty(query))
                {
                    using (OleDbCommand cmd = new OleDbCommand(query, conn))
                    {
                        foreach (string val in values)
                        {
                            cmd.Parameters.Add(new OleDbParameter("?", OleDbType.VarChar)).Value = val;
                        }

                        cmd.ExecuteNonQuery();
                        close();

                        Console.WriteLine($"Cliente adicionado com sucesso.");
                        Console.ReadKey();
                    }

                    return true;
                }
            }
            catch (Exception e)
            {
                Console.WriteLine($"Error: {e.Message}");
                Console.ReadKey();
            }
            return false;
        }
    }
}
