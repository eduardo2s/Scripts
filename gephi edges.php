<?php
$nome_final="C:\wamp64\www\scriptphp\autores.csv";

$lines = file($nome_final, FILE_IGNORE_NEW_LINES);
foreach ($lines as $key)
{
    $vetor = explode(";",$key);
    $tamanho = count($vetor);
    echo $key."</br>";
    echo $tamanho."</br>";
    $nomes = $vetor[0];
    
    for ($i=0;$i < $tamanho-1; $i++){
        for($j=1;$j<$tamanho;$j++){
            if(strcmp($vetor[$i],$vetor[$j]) != 0)
                echo $vetor[$i]." - ".$vetor[$j]."</br>";
        }
    }
}
?>