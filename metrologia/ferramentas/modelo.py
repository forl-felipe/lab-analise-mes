"""
Modelo de dados do sistema Metrologia (fonte única).

Tudo o que descreve as listas do SharePoint mora aqui: nomes, colunas, tipos,
índices e o significado da coluna Título (Title) de cada lista. Os outros
scripts (carga inicial, script PnP, documentação) leem este arquivo, então uma
mudança feita aqui aparece em todos eles.

Tipos de coluna:
    Text      Linha única de texto
    Note      Várias linhas de texto (texto sem formatação)
    Number    Número
    Date      Data (sem hora)
    DateTime  Data e hora
    Boolean   Sim/Não
    Choice    Escolha (choices = lista de opções)
    User      Pessoa ou Grupo (uma pessoa)
"""

from dataclasses import dataclass, field


@dataclass
class Col:
    nome: str
    tipo: str
    descricao: str = ""
    obrigatoria: bool = False
    padrao: str | None = None
    choices: list[str] = field(default_factory=list)
    indexada: bool = False


@dataclass
class Lista:
    nome: str
    titulo_significa: str
    descricao: str
    colunas: list[Col]
    quem_edita: str
    existente: bool = False  # lista que já existe da versão 1
    titulo_unico: bool = False
    titulo_indexado: bool = False

    def coluna(self, nome: str) -> Col:
        return next(c for c in self.colunas if c.nome == nome)


DIAS = [("Seg", "segunda"), ("Ter", "terça"), ("Qua", "quarta"), ("Qui", "quinta"),
        ("Sex", "sexta"), ("Sab", "sábado"), ("Dom", "domingo")]
COLUNAS_AGENDA = [f"{d}{t}" for d, _ in DIAS for t in ("D", "N")]

PERFIS = ["Laboratorista", "Instrumentação"]
LETRAS = ["A", "B", "C", "D", "ADM"]
ENTRADAS = ["Valor", "Tempo", "Comparativo", "ComparativoReferencia", "Massa", "Checklist"]
TIPOS_REGISTRO = ["Calibração", "Verificação", "Comparativo"]

LISTAS: list[Lista] = [
    Lista(
        nome="Acessos",
        titulo_significa="Nome da pessoa (só para leitura humana)",
        descricao="Define quem é Laboratorista e quem é Instrumentação. Quem não está aqui (ou está inativo) entra como Colaborador, somente consulta.",
        quem_edita="Gestor do sistema (dono do site)",
        colunas=[
            Col("Pessoa", "User", "A pessoa, escolhida pelo seletor de pessoas (evita erro de digitação de e-mail)", obrigatoria=True),
            Col("Perfil", "Choice", "Laboratorista: lança aferições. Instrumentação: registra calibrações.", obrigatoria=True, choices=PERFIS),
            Col("Letra", "Choice", "Letra da escala da pessoa (A, B, C, D ou ADM). Vai gravada em cada aferição.", choices=LETRAS),
            Col("Ativo", "Boolean", "Desmarque para tirar o acesso sem apagar o histórico", padrao="1"),
            Col("Observacao", "Text", "Livre"),
        ],
    ),
    Lista(
        nome="Equipamentos",
        existente=True,
        titulo_significa="TAG do equipamento (única; é o nome da imagem do QR Code: TAG.png)",
        descricao="Cadastro único de equipamentos: calibração (instrumentação) e situação da última aferição (laboratório).",
        quem_edita="Instrumentação (calibração) e gestor (cadastro). A situação da aferição é atualizada pelo fluxo F1.",
        titulo_unico=True, titulo_indexado=True,
        colunas=[
            Col("Descricao", "Text", "Ex.: Tambor de abrasão"),
            Col("TipoEquipamento", "Text", "Categoria: Balança, Tambor, Peneira, Prensa..."),
            Col("FabricanteModelo", "Text"),
            Col("NumeroSerie", "Text"),
            Col("Local", "Text", "Área ou bancada"),
            Col("Grupo", "Text", "Agrupador. Para peneiras: o peneirador (Circulares, Peneirador 05...)"),
            Col("Conjunto", "Text", "Peneiras: TAG do peneirador onde a peneira está instalada agora (vazio = fora de uso / em estoque)", indexada=True),
            Col("EnsaioAfericao", "Text", "Código do ensaio de aferição deste equipamento (ex.: TAM). Vazio = não é aferido. Mais de um: separar com ponto e vírgula."),
            Col("ControlaCalibracao", "Boolean", "Sim = tem calibração periódica controlada", padrao="1"),
            Col("PeriodicidadeMeses", "Number", "Intervalo entre calibrações", padrao="12"),
            Col("UltimaCalibracao", "Date"),
            Col("ProximaCalibracao", "Date"),
            Col("UltimoResultado", "Text", "Aprovada ou Recusada"),
            Col("UltimoResponsavel", "Text"),
            Col("UltimoCertificado", "Text", "Número do certificado da última calibração (gravado pelo app). Vai na imagem de consulta"),
            Col("MotivoRecusa", "Note"),
            Col("ManutencaoSolicitada", "Boolean", padrao="0"),
            Col("UltimaAfericao", "Date", "Gravada pelo fluxo F1"),
            Col("UltimaAfericaoResultado", "Text", "Conforme ou Não conforme (fluxo F1)"),
            Col("UltimaAfericaoLancamento", "Text", "Código do lançamento (fluxo F1)"),
            Col("Status", "Text", "Situação calculada pela rotina diária (F2): Liberado / Bloqueado · motivo"),
            Col("ObservacaoCadastro", "Note", "Pendências de cadastro (ex.: TAG a confirmar)"),
            Col("Ativo", "Boolean", padrao="1"),
        ],
    ),
    Lista(
        nome="Ensaios",
        titulo_significa="Código do ensaio (BLA, TAM, ALP, UMI, COM, GRA, T515, PEN, FIS)",
        descricao="Catálogo dos ensaios e agenda semanal (mesma grade da aba Config da planilha).",
        quem_edita="Gestor do sistema",
        titulo_unico=True, titulo_indexado=True,
        colunas=[
            Col("Nome", "Text", "Nome exibido no app", obrigatoria=True),
            Col("NomeBD", "Text", "Nome gravado na coluna Ensaio das aferições (o mesmo da planilha, para o Power BI)", obrigatoria=True),
            Col("Formulario", "Choice", "Medição (valores) ou Checklist (peneiradores)", choices=["Medição", "Checklist"], padrao="Medição"),
            Col("Frequencia", "Text", "Texto exibido no app (ex.: Seg, qua e sáb · noite)"),
            Col("Ordem", "Number", "Ordem na agenda"),
            Col("Instrucoes", "Note", "Aparece no topo do formulário"),
            Col("Ativo", "Boolean", padrao="1"),
        ] + [Col(c, "Boolean", f"Previsto na {next(n for d, n in DIAS if c.startswith(d))} · turno {'07x19 (dia)' if c.endswith('D') else '19x07 (noite)'}", padrao="0") for c in COLUNAS_AGENDA],
    ),
    Lista(
        nome="PontosAfericao",
        titulo_significa="Código do ponto (ex.: TAM-01)",
        descricao="Cada linha é um campo do formulário de um ensaio: qual equipamento, qual parâmetro, que tipo de entrada e qual critério usar. É o equivalente às linhas da aba Lançamento. Nos peneiradores, cada linha é uma POSIÇÃO (peneirador + malha) e a coluna TAG diz qual peneira está nela.",
        quem_edita="Gestor do sistema. Instrumentação troca a peneira de uma posição pelo app",
        titulo_unico=True,
        colunas=[
            Col("Ensaio", "Text", "Código do ensaio (lista Ensaios)", obrigatoria=True, indexada=True),
            Col("Ordem", "Number", "Ordem no formulário"),
            Col("Secao", "Text", "Subtítulo dentro do formulário (ex.: FX -16,0 +12,5 mm, Peneirador 05)"),
            Col("Conjunto", "Text", "Peneiradores: TAG do peneirador desta posição. O QR Code do peneirador abre só as posições dele"),
            Col("TAG", "Text", "TAG do equipamento (lista Equipamentos). Vazio quando não há equipamento cadastrado"),
            Col("EquipamentoBD", "Text", "Texto gravado na coluna Equipamento das aferições (igual à planilha)"),
            Col("Parametro", "Text", obrigatoria=True),
            Col("Unidade", "Text"),
            Col("Entrada", "Choice", "Como o laboratorista informa o resultado", obrigatoria=True, choices=ENTRADAS),
            Col("ChaveLimite", "Text", "Chave do critério na lista Limites. Vazio = informativo"),
            Col("Fator", "Number", "Tempo: número de voltas do tambor. Massa (granulometria): diâmetro médio da fração"),
            Col("RotuloA", "Text", "Rótulo do 1º campo"),
            Col("RotuloB", "Text", "Rótulo do 2º campo (Tempo, Comparativo, Massa)"),
            Col("TipoRegistro", "Choice", "Coluna Tipo das aferições", choices=TIPOS_REGISTRO),
            Col("Ativo", "Boolean", padrao="1"),
        ],
    ),
    Lista(
        nome="Limites",
        titulo_significa="Chave do critério (ex.: TAMB, BLAM, COM16)",
        descricao="Critérios de aceitação. É AQUI que se define nominal, limites e tolerância de cada ensaio. O app usa o valor novo imediatamente.",
        quem_edita="Gestor do sistema",
        titulo_unico=True, titulo_indexado=True,
        colunas=[
            Col("Ensaio", "Text", "Nome do ensaio (como na planilha)"),
            Col("Parametro", "Text"),
            Col("Unidade", "Text"),
            Col("Nominal", "Number", "Valor de referência"),
            Col("LimInferior", "Number", "Se houver limites, Conforme quando LimInferior <= valor <= LimSuperior"),
            Col("LimSuperior", "Number"),
            Col("Tolerancia", "Number", "Sem limites: Conforme quando |valor - referência| <= tolerância"),
            Col("Fonte", "Note", "De onde veio o critério / quem definiu"),
        ],
    ),
    Lista(
        nome="Afericoes",
        titulo_significa="Código do lançamento (ex.: LCP-261024-213501-MA)",
        descricao="Um resultado por linha. Mesmo conteúdo da aba BD_Afericoes (o Power BI lê esta lista).",
        quem_edita="App (laboratoristas). Ninguém edita à mão.",
        titulo_indexado=True,
        colunas=[
            Col("Ensaio", "Text", "Nome do ensaio (NomeBD)", indexada=True),
            Col("TipoRegistro", "Text", "Calibração / Verificação / Comparativo"),
            Col("DataEnsaio", "Date", indexada=True),
            Col("Equipamento", "Text", "Como na planilha (ex.: analisador x Estufa)"),
            Col("TAG", "Text", "TAG do equipamento, para ligar com a lista Equipamentos", indexada=True),
            Col("Parametro", "Text"),
            Col("Unidade", "Text"),
            Col("Valor", "Number"),
            Col("Referencia", "Number"),
            Col("Diferenca", "Number"),
            Col("LimInferior", "Number", "Critério vigente no momento do registro"),
            Col("LimSuperior", "Number"),
            Col("Tolerancia", "Number"),
            Col("Resultado", "Text", "Conforme / Não conforme / Informativo"),
            Col("Responsavel", "Text"),
            Col("ResponsavelEmail", "Text"),
            Col("Letra", "Text"),
            Col("Turno", "Text", "Turno do ensaio: 07x19 ou 19x07"),
            Col("Observacao", "Note"),
            Col("Causa", "Text", "Causa provável (quando não conforme)"),
            Col("ItensNaoOK", "Text", "Peneiradores: itens marcados NÃO OK"),
            Col("RegistradoEm", "DateTime"),
            Col("TurnoRegistro", "Text"),
            Col("ForaDoTurno", "Text", "Sim quando lançado depois do turno do ensaio"),
            Col("Origem", "Text", "App ou Planilha (importado)"),
            Col("ModoLeitura", "Text", "Como o equipamento foi aberto no app: QR (scanner do app), Digitado (TAG digitada, ex.: no computador) ou Link (botão de um e-mail)"),
        ],
    ),
    Lista(
        nome="Lancamentos",
        titulo_significa="Código do lançamento (o mesmo das aferições)",
        descricao="Um resumo por envio (substitui a aba Historico). Dispara o fluxo F1.",
        quem_edita="App (laboratoristas)",
        titulo_unico=True, titulo_indexado=True,
        colunas=[
            Col("Ensaio", "Text", "Nome do ensaio (NomeBD)", indexada=True),
            Col("DataEnsaio", "Date", indexada=True),
            Col("Turno", "Text"),
            Col("Letra", "Text"),
            Col("Responsavel", "Text"),
            Col("ResponsavelEmail", "Text"),
            Col("QtdResultados", "Number"),
            Col("QtdNaoConformes", "Number"),
            Col("TAGsNaoConformes", "Note"),
            Col("Causa", "Text"),
            Col("Observacao", "Note"),
            Col("RegistradoEm", "DateTime"),
            Col("ForaDoTurno", "Text"),
            Col("Situacao", "Text", "Ativo ou Desfeito", padrao="Ativo"),
            Col("Origem", "Text"),
            Col("ModoLeitura", "Text", "Como o equipamento foi aberto no app: QR (scanner do app), Digitado (TAG digitada, ex.: no computador) ou Link (botão de um e-mail)"),
        ],
    ),
    Lista(
        nome="Calibracoes",
        existente=True,
        titulo_significa="TAG do equipamento",
        descricao="Histórico de calibrações (já existe da versão 1).",
        quem_edita="App (instrumentação)",
        colunas=[
            Col("EquipamentoID", "Number", indexada=True),
            Col("DataCalibracao", "Date"),
            Col("Responsavel", "Text"),
            Col("Local", "Text"),
            Col("Resultado", "Text"),
            Col("MotivoRecusa", "Note"),
            Col("ManutencaoSolicitada", "Boolean", padrao="0"),
            Col("ProximaCalibracao", "Date"),
            Col("Certificado", "Text"),
            Col("Observacoes", "Note", "Observação do técnico. As 3 últimas aparecem para a instrumentação na ficha do equipamento"),
            Col("ModoLeitura", "Text", "Como o equipamento foi aberto no app: QR (scanner do app), Digitado (TAG digitada, ex.: no computador) ou Link (botão de um e-mail)"),
        ],
    ),
    Lista(
        nome="Configuracoes",
        titulo_significa="Chave da configuração",
        descricao="Parâmetros gerais: escala das letras, horário dos turnos, massas de ensaio, e-mails dos relatórios.",
        quem_edita="Gestor do sistema",
        titulo_unico=True,
        colunas=[
            Col("Valor", "Text", obrigatoria=True),
            Col("Descricao", "Note"),
        ],
    ),
]


# Biblioteca de documentos com uma imagem por equipamento (TAG.png). O QR Code da
# etiqueta aponta para essa imagem: a câmera do celular abre a imagem; o scanner do
# app lê o mesmo endereço e tira a TAG do nome do arquivo.
BIBLIOTECA_IMAGENS = {
    "nome": "ConsultaQR",
    # Exibição da lista Equipamentos na mesma ordem das colunas do gerador de imagens:
    # copiar as linhas no modo de grade e colar no gerador.
    "exibicao": "Imagens de consulta",
    "campos_exibicao": ["LinkTitle", "Descricao", "Local", "Status", "UltimaCalibracao", "ProximaCalibracao", "UltimoCertificado", "PeriodicidadeMeses"],
    "descricao": "Uma imagem por equipamento (TAG.png), aberta pela câmera do celular ao ler o QR Code. Para atualizar: Carregar o arquivo com o mesmo nome > Substituir.",
}


def lista(nome: str) -> Lista:
    return next(l for l in LISTAS if l.nome == nome)


# Grupos do SharePoint (permissões reais). O app mostra os botões pelo perfil
# da lista Acessos; os grupos garantem que ninguém grave fora do app.
GRUPOS = {
    "Metrologia - Laboratoristas": {"Afericoes": "Contribute", "Lancamentos": "Contribute"},
    # Instrumentação também edita PontosAfericao: é ela quem troca a peneira instalada em cada posição
    "Metrologia - Instrumentação": {"Calibracoes": "Contribute", "Equipamentos": "Contribute", "PontosAfericao": "Contribute"},
}
