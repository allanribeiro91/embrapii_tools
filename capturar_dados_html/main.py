from bs4 import BeautifulSoup

def extract_text_between_tags(html, start_tag, end_tag):
    # Parse o HTML usando BeautifulSoup
    soup = BeautifulSoup(html, 'html.parser')
    
    # Extrair o texto dentro da tag especificada
    start_tag_name = start_tag.split()[0][1:]  # Obtém o nome da tag de abertura (ex: "h3")
    class_name = start_tag.split('"')[1]      # Obtém o nome da classe (ex: "box-title")
    
    # Encontrar todas as tags com o nome e classe especificada
    tags = soup.find_all(start_tag_name, class_=class_name)
    
    # Retornar o texto encontrado em todas as tags, ou uma mensagem se não encontrado
    return [tag.text.strip() for tag in tags] if tags else ["Tag não encontrada"]

# Exemplo de uso
with open('arquivo.html', 'r', encoding='utf-8') as file:
    html = file.read()

start_tag = '<h3 class="box-title">'
end_tag = '</h3>'

textos_extraidos = extract_text_between_tags(html, start_tag, end_tag)
for texto in textos_extraidos:
    print(texto)
