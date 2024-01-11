<?php

if ($_SERVER["REQUEST_METHOD"] == "POST") {
    // Coletar os dados do formulário
    $nome = $_POST["nome"];
    $email = $_POST["email"];
    $mensagem = $_POST["mensagem"];

    // Exibir os dados (substitua isso pelo processamento real)
    echo "Nome: $nome <br>";
    echo "E-mail: $email <br>";
    echo "Mensagem: $mensagem <br>";
} else {
    // Caso alguém acesse este script diretamente sem usar o formulário
    echo "Acesso inválido!";
}

?>
