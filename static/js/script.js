// JavaScript para Otimização de CDs - NexusNode
// Sistema de Otimização de CDs - NexusNode
// Versão: Formato Matriz Única (Solver Excel)

// Função para mostrar passo 2 após download
function mostrarPasso2() {
    // Pega os valores digitados para garantir que não estão vazios
    const cds = document.getElementById('num_cds').value;
    const clientes = document.getElementById('num_clientes').value;

    if (cds > 0 && clientes > 0) {
        // Dá um pequeno atraso (meio segundo) para o download iniciar antes de mostrar a nova tela
        setTimeout(function() {
            const step2 = document.getElementById('step2');
            step2.classList.remove('hidden');
            
            // Adiciona animação suave
            setTimeout(() => {
                step2.style.opacity = '1';
                step2.style.transform = 'translateY(0)';
            }, 50);
            
            // Rola suavemente para a segunda etapa
            step2.scrollIntoView({ behavior: 'smooth', block: 'center' });
        }, 500);
    }
}

// Função para voltar ao passo 1
function backToStep1() {
    document.getElementById('step2').classList.add('hidden');
    document.getElementById('step1').classList.remove('hidden');
    document.getElementById('step2Indicator').classList.remove('active');
    document.getElementById('step1Indicator').classList.add('active');
}

// File upload
const dropZone = document.getElementById('dropZone');
const fileInput = document.getElementById('arquivo_planilha');

if (dropZone && fileInput) {
    dropZone.addEventListener('click', () => fileInput.click());

    dropZone.addEventListener('dragover', (e) => {
        e.preventDefault();
        dropZone.classList.add('dragover');
    });

    dropZone.addEventListener('dragleave', () => {
        dropZone.classList.remove('dragover');
    });

    dropZone.addEventListener('drop', (e) => {
        e.preventDefault();
        dropZone.classList.remove('dragover');
        
        const files = e.dataTransfer.files;
        if (files.length > 0) {
            handleFileSelect(files[0]);
        }
    });

    fileInput.addEventListener('change', function(e) {
        if (e.target.files[0]) {
            handleFileSelect(e.target.files[0]);
        }
    });
}

function handleFileSelect(file) {
    if (file.type === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet') {
        const uploadBtn = document.getElementById('uploadBtn');
        
        // Visual feedback
        dropZone.classList.add('has-file');
        document.querySelector('.upload-content p').textContent = file.name;
        document.querySelector('.upload-content span').textContent = 'Arquivo selecionado';
        
        // Habilitar botão de upload
        uploadBtn.disabled = false;
        uploadBtn.innerHTML = '<span class="material-symbols-outlined">play_arrow</span> Otimizar';
        
        // Mostrar notificação
        showNotification('Arquivo selecionado com sucesso!', 'success');
    } else {
        alert('Por favor, selecione um arquivo .xlsx válido.');
        showNotification('Formato de arquivo inválido!', 'error');
    }
}

// Configuração do formulário de upload - REMOVIDO O PREVENT DEFAULT
document.addEventListener('DOMContentLoaded', function() {
    const uploadForm = document.getElementById('uploadForm');
    if (uploadForm) {
        uploadForm.addEventListener('submit', function(e) {
            // NÃO impedir o comportamento padrão do formulário
            // Deixe o formulário submeter normalmente para o Python
            
            const fileInput = document.getElementById('arquivo_planilha');
            const uploadBtn = document.getElementById('uploadBtn');
            
            if (fileInput.files.length > 0) {
                // Mostrar feedback de processamento
                const originalText = uploadBtn.innerHTML;
                uploadBtn.innerHTML = '<span class="material-symbols-outlined">hourglass_empty</span> Processando...';
                uploadBtn.disabled = true;
                
                // Mostrar notificação
                showNotification('Processando otimização...', 'info');
            } else {
                showNotification('Por favor, selecione um arquivo!', 'error');
                e.preventDefault(); // Só previne se não houver arquivo
            }
        });
    }
});

function showNotification(message, type = 'info') {
    // Criar elemento de notificação
    const notification = document.createElement('div');
    notification.className = `notification notification-${type}`;
    notification.innerHTML = `
        <span class="notification-icon">${type === 'success' ? '✅' : type === 'error' ? '❌' : 'ℹ️'}</span>
        <span class="notification-message">${message}</span>
        <button class="notification-close" onclick="this.parentElement.remove()">×</button>
    `;
    
    // Estilos inline para a notificação
    notification.style.cssText = `
        position: fixed;
        top: 20px;
        right: 20px;
        background: ${type === 'success' ? '#10b981' : type === 'error' ? '#ef4444' : '#3b82f6'};
        color: white;
        padding: 1rem 1.5rem;
        border-radius: 0.75rem;
        box-shadow: 0 8px 25px rgba(0, 0, 0, 0.15);
        display: flex;
        align-items: center;
        gap: 1rem;
        z-index: 10000;
        animation: slideIn 0.3s ease;
        max-width: 400px;
        border-left: 4px solid ${type === 'success' ? '#059669' : type === 'error' ? '#dc2626' : '#1d4ed8'};
    `;
    
    document.body.appendChild(notification);
    
    // Remover após 4 segundos
    setTimeout(() => {
        if (notification.parentElement) {
            notification.remove();
        }
    }, 4000);
}

// Adicionar CSS para animações
const style = document.createElement('style');
style.textContent = `
    @keyframes slideIn {
        from {
            opacity: 0;
            transform: translateX(100%);
        }
        to {
            opacity: 1;
            transform: translateX(0);
        }
    }
    
    .hidden {
        display: none !important;
    }
    
    .dragover {
        background-color: #f0f9ff !important;
        border-color: #3b82f6 !important;
    }
    
    .has-file {
        background-color: #f0fdf4 !important;
        border-color: #10b981 !important;
    }
`;
document.head.appendChild(style);

// Validação de entrada em tempo real
document.addEventListener('DOMContentLoaded', function() {
    const numCdsInput = document.getElementById('num_cds');
    const numClientesInput = document.getElementById('num_clientes');
    
    if (numCdsInput) {
        numCdsInput.addEventListener('input', function() {
            const value = parseInt(this.value);
            const helpText = this.parentElement.querySelector('small');
            
            if (value < 1 || value > 50) {
                this.style.borderColor = '#ef4444';
                if (helpText) {
                    helpText.style.color = '#ef4444';
                    helpText.textContent = 'Valor deve estar entre 1 e 50';
                }
            } else {
                this.style.borderColor = '#10b981';
                if (helpText) {
                    helpText.style.color = '#6b7280';
                    helpText.textContent = 'Número de centros de distribuição disponíveis para instalação';
                }
            }
        });
    }
    
    if (numClientesInput) {
        numClientesInput.addEventListener('input', function() {
            const value = parseInt(this.value);
            const helpText = this.parentElement.querySelector('small');
            
            if (value < 1 || value > 200) {
                this.style.borderColor = '#ef4444';
                if (helpText) {
                    helpText.style.color = '#ef4444';
                    helpText.textContent = 'Valor deve estar entre 1 e 200';
                }
            } else {
                this.style.borderColor = '#10b981';
                if (helpText) {
                    helpText.style.color = '#6b7280';
                    helpText.textContent = 'Número de pontos de entrega que precisam ser atendidos';
                }
            }
        });
    }
});

// Melhora a acessibilidade
document.addEventListener('DOMContentLoaded', function() {
    // Adiciona aria-labels para melhor acessibilidade
    const buttons = document.querySelectorAll('.btn');
    buttons.forEach(btn => {
        if (!btn.getAttribute('aria-label')) {
            btn.setAttribute('aria-label', btn.textContent.trim());
        }
    });
    
    // Adiciona role para as seções
    const sections = document.querySelectorAll('.section, .card');
    sections.forEach(section => {
        section.setAttribute('role', 'region');
    });
    
    // Melhora a navegação por teclado
    document.addEventListener('keydown', function(e) {
        // Enter para submeter formulários
        if (e.key === 'Enter' && e.target.tagName === 'INPUT') {
            const form = e.target.closest('form');
            if (form) {
                form.submit();
            }
        }
    });
});
