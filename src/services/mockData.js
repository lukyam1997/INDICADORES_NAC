const mockSummary = {
  ambulatorial: {
    atendimentos: 1824,
    tempoMedio: 43,
    taxaResolutividade: 0.86,
    retornos: 214,
    tendencia: 0.04,
  },
  cirurgico: {
    procedimentos: 348,
    ocupacao: 0.78,
    mediaPermanencia: 4.2,
    tempoEspera: 16,
    tendencia: -0.03,
  },
};

const mockSeries = {
  ambulatorial: [
    { mes: "Jan", total: 1320, meta: 1200 },
    { mes: "Fev", total: 1480, meta: 1300 },
    { mes: "Mar", total: 1586, meta: 1350 },
    { mes: "Abr", total: 1624, meta: 1400 },
    { mes: "Mai", total: 1702, meta: 1450 },
    { mes: "Jun", total: 1824, meta: 1500 },
  ],
  cirurgico: [
    { mes: "Jan", total: 246, meta: 220 },
    { mes: "Fev", total: 258, meta: 240 },
    { mes: "Mar", total: 292, meta: 250 },
    { mes: "Abr", total: 310, meta: 260 },
    { mes: "Mai", total: 328, meta: 280 },
    { mes: "Jun", total: 348, meta: 300 },
  ],
};

const mockRecords = {
  ambulatorial: [
    {
      paciente: "Maria Silva",
      especialidade: "Cardiologia",
      profissional: "Dr. João Alves",
      data: "2024-06-06",
      status: "Concluído",
    },
    {
      paciente: "Bruno Lima",
      especialidade: "Endocrinologia",
      profissional: "Dra. Fernanda Rocha",
      data: "2024-06-06",
      status: "Em andamento",
    },
  ],
  cirurgico: [
    {
      paciente: "Ana Souza",
      procedimento: "Colecistectomia",
      equipe: "Equipe 3",
      data: "2024-06-05",
      status: "Agendado",
    },
    {
      paciente: "Carlos Mendes",
      procedimento: "Herniorrafia",
      equipe: "Equipe 1",
      data: "2024-06-05",
      status: "Recuperação",
    },
  ],
};

export { mockSummary, mockSeries, mockRecords };
