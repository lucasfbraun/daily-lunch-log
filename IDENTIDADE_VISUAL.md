# 🎨 Identidade Visual - Dashboard de Licenciamento Microsoft

## 📋 Paleta de Cores Material Design

### Cores Principais

| Função / Uso | Cor | Hexadecimal | RGB | Uso no Dashboard |
|--------------|-----|-------------|-----|------------------|
| **Dark 1 (Preto)** | ![#000000](https://via.placeholder.com/15/000000/000000?text=+) | `#000000` | `rgb(0, 0, 0)` | Textos de alto contraste |
| **Light 1 (Branco)** | ![#FFFFFF](https://via.placeholder.com/15/FFFFFF/000000?text=+) | `#FFFFFF` | `rgb(255, 255, 255)` | Background de cards, textos em headers |
| **Dark 2 (Cinza Escuro)** | ![#595959](https://via.placeholder.com/15/595959/000000?text=+) | `#595959` | `rgb(89, 89, 89)` | Texto padrão, títulos |
| **Light 2 (Cinza Claro)** | ![#EEEEEE](https://via.placeholder.com/15/EEEEEE/000000?text=+) | `#EEEEEE` | `rgb(238, 238, 238)` | Background geral, separadores |
| **Accent 1 (Azul Google)** | ![#4285F4](https://via.placeholder.com/15/4285F4/000000?text=+) | `#4285F4` | `rgb(66, 133, 244)` | Header, KPI usuários, gráfico empresas, títulos |
| **Accent 2 (Preto Fosco)** | ![#212121](https://via.placeholder.com/15/212121/000000?text=+) | `#212121` | `rgb(33, 33, 33)` | Títulos escuros, hover de botões |
| **Accent 3 (Cinza Azulado)** | ![#78909C](https://via.placeholder.com/15/78909C/000000?text=+) | `#78909C` | `rgb(120, 144, 156)` | KPI empresas, gráfico licenças, scrollbar |
| **Accent 4 (Laranja)** | ![#FFAB40](https://via.placeholder.com/15/FFAB40/000000?text=+) | `#FFAB40` | `rgb(255, 171, 64)` | KPI licenças, gráfico setores, alertas |
| **Accent 5 (Ciano)** | ![#0097A7](https://via.placeholder.com/15/0097A7/000000?text=+) | `#0097A7` | `rgb(0, 151, 167)` | Header gradient, KPI gasto, badges, links |
| **Accent 6 (Amarelo Neon)** | ![#EEFF41](https://via.placeholder.com/15/EEFF41/000000?text=+) | `#EEFF41` | `rgb(238, 255, 65)` | Highlights (uso futuro) |

---

## 🎯 Aplicação das Cores

### Header Principal
```css
background: linear-gradient(135deg, #4285F4 0%, #0097A7 100%);
color: #FFFFFF;
border-radius: 16px;
box-shadow: 0 8px 16px rgba(66, 133, 244, 0.25);
```
- **Gradiente**: Azul Google → Ciano
- **Texto**: Branco
- **Efeito**: Moderno e tecnológico

### KPI Cards

#### 💰 Gasto Total
```css
background: linear-gradient(135deg, #0097A7 0%, #007c8a 100%);
color: #FFFFFF;
```
- **Cor Base**: Accent 5 (Ciano)
- **Mensagem**: Crescimento e controle financeiro

#### 👥 Total de Usuários
```css
background: linear-gradient(135deg, #4285F4 0%, #3367d6 100%);
color: #FFFFFF;
```
- **Cor Base**: Accent 1 (Azul Google)
- **Mensagem**: Tecnologia e confiança

#### 🏢 Empresas
```css
background: linear-gradient(135deg, #78909C 0%, #607D8B 100%);
color: #FFFFFF;
```
- **Cor Base**: Accent 3 (Cinza Azulado)
- **Mensagem**: Profissionalismo e solidez

#### 📋 Licenças
```css
background: linear-gradient(135deg, #FFAB40 0%, #FF9100 100%);
color: #000000;
```
- **Cor Base**: Accent 4 (Laranja)
- **Texto**: Preto (contraste)
- **Mensagem**: Energia e atenção

### Gráficos

#### 💼 Gastos por Empresa
- **Cor**: `#4285F4` (Azul Google - Accent 1)
- **Background**: `#FFFFFF` (Branco)
- **Fonte**: Roboto, sans-serif

#### 🗺️ Distribuição por Estado (Pizza)
- **Paleta**: `['#4285F4', '#0097A7', '#78909C', '#FFAB40', '#212121', '#EEEEEE']`
- **Ordem**: Azul Google → Ciano → Cinza Azulado → Laranja → Preto Fosco → Cinza Claro

#### 🏦 Centros de Custo
- **Cor**: `#0097A7` (Ciano - Accent 5)
- **Background**: `#FFFFFF` (Branco)
- **Fonte**: Roboto, sans-serif

#### 📊 Licenças Mais Usadas
- **Cor**: `#78909C` (Cinza Azulado - Accent 3)
- **Background**: `#FFFFFF` (Branco)
- **Fonte**: Roboto, sans-serif

#### 💳 Modalidade de Licença (Pizza)
- **Paleta**: `['#4285F4', '#0097A7', '#FFAB40', '#78909C', '#212121']`
- **Ordem**: Azul → Ciano → Laranja → Cinza Azulado → Preto

#### 🏢 Gastos por Setor
- **Cor**: `#FFAB40` (Laranja - Accent 4)
- **Background**: `#FFFFFF` (Branco)
- **Fonte**: Roboto, sans-serif

#### 🔄 Distribuição por Fornecedor (Pizza)
- **Paleta**: `['#4285F4', '#0097A7', '#78909C', '#FFAB40', '#212121', '#EEEEEE']`

---

## 🎨 Elementos Interativos

### Botões Principais
```css
background: linear-gradient(135deg, #4285F4 0%, #0097A7 100%);
color: #FFFFFF;
border-radius: 10px;
text-transform: uppercase;
letter-spacing: 0.8px;
```

**Hover**:
```css
background: linear-gradient(135deg, #0097A7 0%, #4285F4 100%);
transform: translateY(-3px);
box-shadow: 0 6px 16px rgba(66, 133, 244, 0.4);
```

### Botões Secundários (Limpar)
```css
background: #595959;
color: #FFFFFF;
```

**Hover**:
```css
background: #212121;
transform: translateY(-3px);
```

### Badges de Licença
```css
background: #0097A7;
color: #FFFFFF;
padding: 10px 18px;
border-radius: 8px;
```

**Hover**:
```css
background: #4285F4;
transform: scale(1.1);
box-shadow: 0 4px 12px rgba(0, 151, 167, 0.5);
```

---

## 📊 Tabelas de Contratos

### Alertas de Vencimento

#### 🔴 Vencido
```css
background-color: #ffebee;
border-left: 5px solid #f44336;
font-weight: 600;
```

#### ⚠️ Vence em breve (≤7 dias)
```css
background-color: #fff8e1;
border-left: 5px solid #FFAB40; /* Laranja */
font-weight: 600;
```

#### 🔵 Atenção (≤30 dias)
```css
background-color: #e1f5fe;
border-left: 5px solid #0097A7; /* Ciano */
```

#### ✅ OK (>30 dias)
```css
background-color: #FFFFFF;
border-left: none;
```

### Header de Tabela
```css
background: linear-gradient(135deg, #4285F4 0%, #0097A7 100%);
color: #FFFFFF;
text-transform: uppercase;
letter-spacing: 0.8px;
```

---

## 🎯 Cards de Conteúdo

### Card Padrão
```css
background: #FFFFFF;
border: none;
border-radius: 16px;
box-shadow: 0 2px 8px rgba(0, 0, 0, 0.06);
transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1);
```

**Hover**:
```css
box-shadow: 0 8px 20px rgba(0, 0, 0, 0.12);
transform: translateY(-2px);
```

### Filtros Section
```css
background: #FFFFFF;
border-top: 4px solid #4285F4; /* Azul Google */
border-radius: 16px;
box-shadow: 0 2px 12px rgba(0, 0, 0, 0.06);
```

### User Cards (Modal)
```css
background: #EEEEEE; /* Light 2 */
border-left: 5px solid #0097A7; /* Ciano */
border-radius: 10px;
```

**Hover**:
```css
background: #e0e0e0;
border-left-color: #4285F4; /* Azul Google */
box-shadow: 0 4px 12px rgba(66, 133, 244, 0.2);
transform: translateX(8px);
```

---

## 🖼️ Composição Visual

### Hierarquia de Cores

1. **Primário**: Azul Google (#4285F4) - Autoridade, tecnologia
2. **Secundário**: Ciano (#0097A7) - Clareza, modernidade
3. **Acento 1**: Cinza Azulado (#78909C) - Profissionalismo
4. **Acento 2**: Laranja (#FFAB40) - Destaque, energia
5. **Neutro Escuro**: Cinza Escuro (#595959) - Texto padrão
6. **Neutro Claro**: Cinza Claro (#EEEEEE) - Background, separação
7. **Base**: Branco (#FFFFFF) - Limpeza, espaço
8. **Destaque Escuro**: Preto Fosco (#212121) - Contraste forte

### Contraste e Acessibilidade

- **Texto em Azul Google/Ciano**: Sempre branco (#FFFFFF)
- **Texto em Laranja**: Preto (#000000) para contraste
- **Texto em Cinza Azulado**: Sempre branco (#FFFFFF)
- **Texto Padrão**: Cinza Escuro (#595959) sobre branco ou cinza claro
- **Títulos Importantes**: Preto Fosco (#212121)

### Proporção de Uso

```
Azul Google:      30% - Headers, gráficos principais, KPIs
Ciano:            25% - Gradientes, badges, links
Branco:           20% - Backgrounds, cards
Cinza Claro:      10% - Background geral
Laranja/Azulado:  10% - Acentos, KPIs, gráficos
Cinza Escuro:     5%  - Texto, títulos
```

---

## 📐 Estilos de Componentes

### Bordas
- **Border Radius Cards**: `16px`
- **Border Radius Botões**: `10px`
- **Border Radius User Cards**: `10px`
- **Border Radius Inputs**: `8px`
- **Border Radius Badges**: `8px`

### Sombras
- **Card Padrão**: `0 2px 8px rgba(0, 0, 0, 0.06)`
- **Card Hover**: `0 8px 20px rgba(0, 0, 0, 0.12)`
- **KPI Hover**: `0 12px 24px rgba(0, 0, 0, 0.15)`
- **Botão Hover**: `0 6px 16px rgba(66, 133, 244, 0.4)`
- **Header**: `0 8px 16px rgba(66, 133, 244, 0.25)`
- **Modal**: `0 12px 32px rgba(0, 0, 0, 0.2)`

### Transições
```css
transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1);
```

**Material Design Easing:**
- Entrada: `cubic-bezier(0.4, 0, 0.2, 1)` - Desaceleração
- Saída: `cubic-bezier(0.4, 0, 1, 1)` - Aceleração
- KPIs: `0.4s` para efeito dramático

---

## 🎨 Variáveis CSS

```css
:root {
    /* Paleta de Cores Material Design */
    --dk1: #000000;
    --lt1: #FFFFFF;
    --dk2: #595959;
    --lt2: #EEEEEE;
    --accent1: #4285F4;
    --accent2: #212121;
    --accent3: #78909C;
    --accent4: #FFAB40;
    --accent5: #0097A7;
    --accent6: #EEFF41;
    --hlink: #0097A7;
    --folHlink: #0097A7;
}
```

---

## 🎯 Tipografia

### Font Family
```css
font-family: 'Roboto', 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
```

### Pesos de Fonte
- **Light**: 300 - Subtítulos, descrições
- **Regular**: 400 - Texto padrão (não usado, preferir 500)
- **Medium**: 500 - Badges, labels
- **Semi-Bold**: 600 - Títulos de seção, labels de filtro
- **Bold**: 700 - Headers, KPIs, títulos principais

### Tamanhos
- **Header H1**: 2.5rem (40px) - Título principal
- **KPI Valor**: 2.8rem (44.8px) - Números grandes
- **Card Title**: 1.3rem (20.8px) - Títulos de cards
- **Graph Title**: 18px - Títulos de gráficos
- **Filter Label**: 0.9rem (14.4px) - Labels de filtros
- **Texto Padrão**: 1rem (16px) - Body text

### Letter Spacing
- **Uppercase Buttons**: 0.8px
- **KPI Labels**: 1.2px
- **Table Headers**: 0.8px
- **Filter Labels**: 0.5px
- **Títulos Grandes**: -0.5px a -1px (ajuste ótico)

---

## ✅ Checklist de Consistência

- [x] Header usa gradiente Azul Google → Ciano
- [x] Todos os gráficos têm background branco
- [x] Texto padrão em Cinza Escuro (#595959)
- [x] KPIs com gradientes das cores Material Design
- [x] Botões principais em gradiente azul-ciano
- [x] Badges em Ciano (#0097A7)
- [x] Alertas de contrato com cores apropriadas
- [x] Cards com sombras sutis Material Design
- [x] Transições com easing cubic-bezier
- [x] Border radius arredondado (10-16px)
- [x] Fonte Roboto em todos os elementos
- [x] Scrollbar personalizada com Accent 3
- [x] Animações de entrada (slideIn)

---

## 🚀 Implementação

### Para aplicar esta identidade visual:

1. **CSS Global**: Todas as cores definidas em variáveis CSS `:root`
2. **Gráficos Plotly**: Cores aplicadas via `marker_color` e `color_discrete_sequence`
3. **Tipografia**: Font family Roboto, com fallback para Segoe UI
4. **Componentes Bootstrap**: Sobrescritos com classes customizadas
5. **Animações**: Easing Material Design para movimento natural
6. **Consistência**: Mesma paleta em todos os elementos

### Arquivos Modificados:
- `dashboard_flask.py` - Funções de criação de gráficos e template HTML com CSS Material Design

---

## 🎨 Recursos Visuais Adicionais

### Scrollbar Personalizada
```css
::-webkit-scrollbar {
    width: 10px;
}

::-webkit-scrollbar-track {
    background: #EEEEEE;
}

::-webkit-scrollbar-thumb {
    background: #78909C;
    border-radius: 5px;
}

::-webkit-scrollbar-thumb:hover {
    background: #0097A7;
}
```

### Animação de Entrada
```css
@keyframes slideIn {
    from {
        opacity: 0;
        transform: translateY(20px);
    }
    to {
        opacity: 1;
        transform: translateY(0);
    }
}

.card-custom, .kpi-card {
    animation: slideIn 0.5s ease-out;
}
```

---

**Criado em**: 15 de Outubro de 2025  
**Versão**: 3.0.0 - Material Design  
**Paleta**: Google Material Design Flexible Palette  
**Inspiração**: Google Material Design Guidelines

