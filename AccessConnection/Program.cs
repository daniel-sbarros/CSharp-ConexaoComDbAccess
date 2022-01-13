// See https://aka.ms/new-console-template for more information
//Console.WriteLine("Hello, World!");
using AccessConnection;
using System.Data.OleDb;
using System.Text;

Conexao conexao = new Conexao();
//conexao.listar("SELECT * FROM Clientes");

StringBuilder opcao = new StringBuilder();

while (opcao.ToString() != "sair")
{
    Console.Clear();

    Console.WriteLine("Digite o número da opção Desejada e Tecle Enter.\n\n");
    Console.WriteLine("1 - Adicionar Clientes.");
    Console.WriteLine("2 - Listar Clientes.");
    Console.WriteLine("3 - Sair.\n");
    
    opcao.Clear();
    opcao.Append(Console.ReadLine());

    switch (opcao.ToString())
    {
        case "1":
            addCliente();
            break;
        case "2":
            Console.Clear();
            listarClientes();
            Console.ReadLine();
            break;
        case "3":
            opcao.Clear();
            opcao.Append("sair");
            break;
        default:
            Console.WriteLine("Opção Inválida.");
            Console.ReadLine();
            break;
    }
}

void addCliente()
{
    if (conexao.open()) {
        Console.Clear();
        Console.WriteLine("Digite os dados do Cliente.\n");
        Console.Write("Nome: ");
        string nome = Console.ReadLine();
        Console.Write("Email: ");
        string email = Console.ReadLine();

        conexao.add("INSERT INTO Clientes (Nome, Email) Values (?, ?)", nome, email);
    }
}

void listarClientes()
{
    conexao.listar("SELECT * FROM Clientes");
}