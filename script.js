// Função para carregar dados do JSON
async function carregarDados() {
    try {
        const response = await fetch('data/database.json');
        const dados = await response.json();
        
        // Atualizar a interface com os dados
        atualizarInterface(dados);
    } catch (error) {
        console.error('Erro ao carregar dados:', error);
    }
}

// Função para atualizar a interface
function atualizarInterface(dados) {
    // Atualizar título
    document.getElementById('titulo-site').textContent = dados.configuracoes.titulo;
    
    // Atualizar produtos
    const container = document.getElementById('produtos-container');
    container.innerHTML = ''; // Limpar container
    
    dados.produtos.forEach(produto => {
        const card = document.createElement('div');
        card.className = 'produto-card';
        card.innerHTML = `
            <h3>${produto.nome}</h3>
            <p>${produto.descricao}</p>
            <span class="preco">R$ ${produto.preco.toFixed(2)}</span>
        `;
        container.appendChild(card);
    });
}

// Carregar dados quando a página carregar
document.addEventListener('DOMContentLoaded', carregarDados);
