import React, { useState, useEffect } from 'react';

interface RapportFlashMatrixProps {
  projectId: string;
  fileId: string;
  currentDossierId: string | null;
  activeSubTab?: string;
}

interface ActionItem {
  id: number;
  action: string;
  actor: string;
  deadline: string;
  status: 'Terminé' | 'En cours' | 'En attente';
  weeklyUpdate: string;
}

interface RiskItem {
  id: number;
  risk: string;
  actionProposal: string;
  actor: string;
}

type WeatherState = '☀️' | '🌥️' | '🌧️';
type TrendState = '↗️' | '➡️' | '↘️';

interface DomainStatus {
  weather: WeatherState;
  trend: TrendState;
  label: string;
  comment: string;
}

const DEFAULT_ACTIONS: ActionItem[] = [
  {
    id: 1,
    action: 'Finaliser le paramétrage et les tests du Grand Livre',
    actor: 'Marianne Anselme',
    deadline: '23/07/2022',
    status: 'Terminé',
    weeklyUpdate: 'Documentation en cours'
  },
  {
    id: 2,
    action: 'Préparer un jeu de test pour les immobilisations',
    actor: 'Gérard Lourte',
    deadline: '24/07/2022',
    status: 'En cours',
    weeklyUpdate: 'Échantillon important de cas spéciaux'
  },
  {
    id: 3,
    action: "Sélection d'un utilisateur clé pour la production",
    actor: 'Georges Clonet',
    deadline: '25/07/2022',
    status: 'En cours',
    weeklyUpdate: 'Marcel Huet est disponible pour nous'
  },
  {
    id: 4,
    action: "Évaluer les besoins d'interfaces pour le dépôt",
    actor: 'Eric Duteil',
    deadline: '26/07/2022',
    status: 'Terminé',
    weeklyUpdate: 'Principalement les systèmes uniques au CMRL'
  },
  {
    id: 5,
    action: 'Préparation du comité de pilotage',
    actor: 'Chantal Davant',
    deadline: '27/07/2022',
    status: 'En attente',
    weeklyUpdate: 'Date du workshop fixée'
  },
  {
    id: 6,
    action: 'Mise à jour du site intranet projet',
    actor: 'Bernard Dupuis',
    deadline: '28/07/2022',
    status: 'En attente',
    weeklyUpdate: 'Préparation terminée, en attente de la revue du sponsor'
  }
];

const DEFAULT_RISKS: RiskItem[] = [
  {
    id: 1,
    risk: 'Qualité relativement médiocre des données',
    actionProposal: "Autoriser une augmentation du temps assigné au projet, voire accepter des heures supplémentaires",
    actor: 'Chantal Davant / Olivier Robespierre'
  },
  {
    id: 2,
    risk: "Absentéisme plus important que d'habitude",
    actionProposal: 'Identifier des ressources expérimentées qui pourraient remplacer les absents',
    actor: 'Daniel Château'
  },
  {
    id: 3,
    risk: 'Infrastructure un peu juste',
    actionProposal: 'Mise en place du plan de contingence prévu dans le registre des risques',
    actor: 'Daniel Château'
  },
  {
    id: 4,
    risk: 'Nombreuses questions venant des syndicats',
    actionProposal: "Terminer le plan de communication et le matériel prévu à cet effet",
    actor: 'Bernard Dupuis'
  },
  {
    id: 5,
    risk: 'Nombreuses questions des utilisateurs impactés par le projet',
    actionProposal: "Jacques Meunier travaillera avec Bernard Dupuis et coordonnera les aspects d'intégration au CMRL",
    actor: 'Jacques Meunier / Bernard Dupuis'
  }
];

export const RapportFlashMatrix: React.FC<RapportFlashMatrixProps> = ({
  projectId,
  fileId,
  currentDossierId,
  activeSubTab = 'rapport'
}) => {
  const dossierSubKey = currentDossierId ? `-${currentDossierId}` : '';
  const storageKey = `faab-rapport-flash-matrix-${projectId}-${fileId}${dossierSubKey}`;

  // Domain Weather & Trends
  const [budgetStatus, setBudgetStatus] = useState<DomainStatus>(() => {
    const saved = localStorage.getItem(`${storageKey}-budget`);
    return saved ? JSON.parse(saved) : { weather: '☀️', trend: '➡️', label: 'OPTIMAL', comment: 'Consommation conforme aux prévisions' };
  });

  const [planningStatus, setPlanningStatus] = useState<DomainStatus>(() => {
    const saved = localStorage.getItem(`${storageKey}-planning`);
    return saved ? JSON.parse(saved) : { weather: '🌥️', trend: '↗️', label: 'VIGILANCE', comment: 'Léger décalage sur les tests d’intégration' };
  });

  const [scopeStatus, setScopeStatus] = useState<DomainStatus>(() => {
    const saved = localStorage.getItem(`${storageKey}-scope`);
    return saved ? JSON.parse(saved) : { weather: '☀️', trend: '➡️', label: 'CONFORME', comment: 'Aucun changement majeur hors périmètre' };
  });

  // Actions Table
  const [actions, setActions] = useState<ActionItem[]>(() => {
    const saved = localStorage.getItem(`${storageKey}-actions`);
    return saved ? JSON.parse(saved) : DEFAULT_ACTIONS;
  });

  // Important Info
  const [importantInfo, setImportantInfo] = useState<string>(() => {
    const saved = localStorage.getItem(`${storageKey}-info`);
    return saved || "• Bonne avancée du paramétrage dans toutes les équipes. Assez peu de corrections après les tests unitaires.\n• Les premières interfaces logistiques sont disponibles.\n• Les définitions des interfaces de production sont désormais terminées et communiquées aux développeurs.\n• Les corrections de données ont débuté. Le travail sera minutieux.";
  });

  // Expected Decisions
  const [expectedDecisions, setExpectedDecisions] = useState<string>(() => {
    const saved = localStorage.getItem(`${storageKey}-decisions`);
    return saved || "• Autorisation d'augmenter le temps assigné à la préparation et la correction des données.\n• Validation des éventuelles heures supplémentaires pour l'équipe Master Data.\n• Arbitrage sur la date définitive du comité de pilotage.";
  });

  // Risks Table
  const [risks, setRisks] = useState<RiskItem[]>(() => {
    const saved = localStorage.getItem(`${storageKey}-risks`);
    return saved ? JSON.parse(saved) : DEFAULT_RISKS;
  });

  // Persistence
  useEffect(() => {
    localStorage.setItem(`${storageKey}-budget`, JSON.stringify(budgetStatus));
  }, [budgetStatus, storageKey]);

  useEffect(() => {
    localStorage.setItem(`${storageKey}-planning`, JSON.stringify(planningStatus));
  }, [planningStatus, storageKey]);

  useEffect(() => {
    localStorage.setItem(`${storageKey}-scope`, JSON.stringify(scopeStatus));
  }, [scopeStatus, storageKey]);

  useEffect(() => {
    localStorage.setItem(`${storageKey}-actions`, JSON.stringify(actions));
  }, [actions, storageKey]);

  useEffect(() => {
    localStorage.setItem(`${storageKey}-info`, importantInfo);
  }, [importantInfo, storageKey]);

  useEffect(() => {
    localStorage.setItem(`${storageKey}-decisions`, expectedDecisions);
  }, [expectedDecisions, storageKey]);

  useEffect(() => {
    localStorage.setItem(`${storageKey}-risks`, JSON.stringify(risks));
  }, [risks, storageKey]);

  // Weather cycler
  const cycleWeather = (current: WeatherState): WeatherState => {
    if (current === '☀️') return '🌥️';
    if (current === '🌥️') return '🌧️';
    return '☀️';
  };

  // Trend cycler
  const cycleTrend = (current: TrendState): TrendState => {
    if (current === '➡️') return '↗️';
    if (current === '↗️') return '↘️';
    return '➡️';
  };

  // Status cycler
  const cycleActionStatus = (status: 'Terminé' | 'En cours' | 'En attente'): 'Terminé' | 'En cours' | 'En attente' => {
    if (status === 'Terminé') return 'En cours';
    if (status === 'En cours') return 'En attente';
    return 'Terminé';
  };

  const getStatusBadge = (status: 'Terminé' | 'En cours' | 'En attente') => {
    if (status === 'Terminé') {
      return { bg: 'rgba(16, 185, 129, 0.15)', color: '#10b981', label: 'Terminé ✅' };
    }
    if (status === 'En cours') {
      return { bg: 'rgba(59, 130, 246, 0.15)', color: '#3b82f6', label: 'En cours 🔵' };
    }
    return { bg: 'rgba(245, 158, 11, 0.15)', color: '#f59e0b', label: 'En attente ⏳' };
  };

  return (
    <div style={{ flex: 1, display: 'flex', flexDirection: 'column', background: 'var(--main-bg)', color: 'var(--main-text)', overflowY: 'auto' }}>
      {/* Top Header */}
      <div style={{ padding: '1.25rem 2rem', background: 'var(--main-surface)', borderBottom: '1px solid var(--main-border)', display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
        <div>
          <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
            <span style={{ fontSize: '1.3rem' }}>⚡</span>
            <h1 style={{ fontSize: '1.25rem', fontWeight: 900, margin: 0, color: 'var(--main-text)' }}>
              30 — Rapport Flash Hebdomadaire
            </h1>
            <span style={{ padding: '4px 10px', background: 'rgba(16, 185, 129, 0.12)', color: '#10b981', borderRadius: '20px', fontSize: '0.72rem', fontWeight: 800 }}>
              EXCEL 30 DYNAMIQUE
            </span>
          </div>
          <p style={{ margin: '4px 0 0', fontSize: '0.8rem', color: 'var(--text-muted)' }}>
            Kit Chef de Projet • Synthèse exécutive : Météo & Tendances, Actions clés, Décisions attendues et Risques
          </p>
        </div>
      </div>

      <div style={{ padding: '2rem', display: 'flex', flexDirection: 'column', gap: '2rem' }}>
        {/* 1. SECTION MÉTÉO ET TENDANCE */}
        <div style={{ background: 'var(--main-surface)', padding: '2rem', borderRadius: '20px', border: '1px solid var(--main-border)', boxShadow: 'var(--card-shadow)' }}>
          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '1.5rem' }}>
            <h3 style={{ margin: 0, fontSize: '1rem', fontWeight: 900, color: '#10b981', textTransform: 'uppercase', letterSpacing: '0.05em' }}>
              🌤️ Indicateurs de Performance & Tendance (Météo Projet)
            </h3>
            <span style={{ fontSize: '0.75rem', color: 'var(--text-muted)', fontWeight: 600 }}>
              💡 Cliquez sur les icônes pour faire évoluer la météo ou la tendance
            </span>
          </div>

          <div style={{ display: 'grid', gridTemplateColumns: 'repeat(3, 1fr)', gap: '1.5rem' }}>
            {/* BUDGET */}
            <div style={{ background: 'var(--main-bg)', padding: '1.5rem', borderRadius: '16px', border: '1px solid var(--main-border)', textAlign: 'center' }}>
              <div style={{ fontSize: '0.75rem', fontWeight: 800, color: 'var(--text-muted)', textTransform: 'uppercase', marginBottom: '8px' }}>
                BUDGET
              </div>
              <div style={{ display: 'flex', justifyContent: 'center', alignItems: 'center', gap: '16px', margin: '8px 0' }}>
                <button
                  onClick={() => setBudgetStatus({ ...budgetStatus, weather: cycleWeather(budgetStatus.weather) })}
                  title="Changer la météo"
                  style={{ fontSize: '2.5rem', background: 'transparent', border: 'none', cursor: 'pointer', transition: 'transform 0.15s ease' }}
                  onMouseEnter={(e) => (e.currentTarget.style.transform = 'scale(1.15)')}
                  onMouseLeave={(e) => (e.currentTarget.style.transform = 'scale(1)')}
                >
                  {budgetStatus.weather}
                </button>
                <button
                  onClick={() => setBudgetStatus({ ...budgetStatus, trend: cycleTrend(budgetStatus.trend) })}
                  title="Changer la tendance"
                  style={{ fontSize: '1.8rem', background: 'transparent', border: 'none', cursor: 'pointer', transition: 'transform 0.15s ease' }}
                  onMouseEnter={(e) => (e.currentTarget.style.transform = 'scale(1.15)')}
                  onMouseLeave={(e) => (e.currentTarget.style.transform = 'scale(1)')}
                >
                  {budgetStatus.trend}
                </button>
              </div>
              <input
                type="text"
                value={budgetStatus.comment}
                onChange={(e) => setBudgetStatus({ ...budgetStatus, comment: e.target.value })}
                placeholder="Commentaire de situation..."
                style={{ width: '90%', padding: '6px 10px', borderRadius: '6px', border: '1px solid var(--main-border)', background: 'var(--main-surface)', color: 'var(--main-text)', fontSize: '0.78rem', textAlign: 'center' }}
              />
            </div>

            {/* PLANNING */}
            <div style={{ background: 'var(--main-bg)', padding: '1.5rem', borderRadius: '16px', border: '1px solid var(--main-border)', textAlign: 'center' }}>
              <div style={{ fontSize: '0.75rem', fontWeight: 800, color: 'var(--text-muted)', textTransform: 'uppercase', marginBottom: '8px' }}>
                PLANNING & DÉLAIS
              </div>
              <div style={{ display: 'flex', justifyContent: 'center', alignItems: 'center', gap: '16px', margin: '8px 0' }}>
                <button
                  onClick={() => setPlanningStatus({ ...planningStatus, weather: cycleWeather(planningStatus.weather) })}
                  title="Changer la météo"
                  style={{ fontSize: '2.5rem', background: 'transparent', border: 'none', cursor: 'pointer', transition: 'transform 0.15s ease' }}
                  onMouseEnter={(e) => (e.currentTarget.style.transform = 'scale(1.15)')}
                  onMouseLeave={(e) => (e.currentTarget.style.transform = 'scale(1)')}
                >
                  {planningStatus.weather}
                </button>
                <button
                  onClick={() => setPlanningStatus({ ...planningStatus, trend: cycleTrend(planningStatus.trend) })}
                  title="Changer la tendance"
                  style={{ fontSize: '1.8rem', background: 'transparent', border: 'none', cursor: 'pointer', transition: 'transform 0.15s ease' }}
                  onMouseEnter={(e) => (e.currentTarget.style.transform = 'scale(1.15)')}
                  onMouseLeave={(e) => (e.currentTarget.style.transform = 'scale(1)')}
                >
                  {planningStatus.trend}
                </button>
              </div>
              <input
                type="text"
                value={planningStatus.comment}
                onChange={(e) => setPlanningStatus({ ...planningStatus, comment: e.target.value })}
                placeholder="Commentaire de situation..."
                style={{ width: '90%', padding: '6px 10px', borderRadius: '6px', border: '1px solid var(--main-border)', background: 'var(--main-surface)', color: 'var(--main-text)', fontSize: '0.78rem', textAlign: 'center' }}
              />
            </div>

            {/* PÉRIMÈTRE */}
            <div style={{ background: 'var(--main-bg)', padding: '1.5rem', borderRadius: '16px', border: '1px solid var(--main-border)', textAlign: 'center' }}>
              <div style={{ fontSize: '0.75rem', fontWeight: 800, color: 'var(--text-muted)', textTransform: 'uppercase', marginBottom: '8px' }}>
                PÉRIMÈTRE & QUALITÉ
              </div>
              <div style={{ display: 'flex', justifyContent: 'center', alignItems: 'center', gap: '16px', margin: '8px 0' }}>
                <button
                  onClick={() => setScopeStatus({ ...scopeStatus, weather: cycleWeather(scopeStatus.weather) })}
                  title="Changer la météo"
                  style={{ fontSize: '2.5rem', background: 'transparent', border: 'none', cursor: 'pointer', transition: 'transform 0.15s ease' }}
                  onMouseEnter={(e) => (e.currentTarget.style.transform = 'scale(1.15)')}
                  onMouseLeave={(e) => (e.currentTarget.style.transform = 'scale(1)')}
                >
                  {scopeStatus.weather}
                </button>
                <button
                  onClick={() => setScopeStatus({ ...scopeStatus, trend: cycleTrend(scopeStatus.trend) })}
                  title="Changer la tendance"
                  style={{ fontSize: '1.8rem', background: 'transparent', border: 'none', cursor: 'pointer', transition: 'transform 0.15s ease' }}
                  onMouseEnter={(e) => (e.currentTarget.style.transform = 'scale(1.15)')}
                  onMouseLeave={(e) => (e.currentTarget.style.transform = 'scale(1)')}
                >
                  {scopeStatus.trend}
                </button>
              </div>
              <input
                type="text"
                value={scopeStatus.comment}
                onChange={(e) => setScopeStatus({ ...scopeStatus, comment: e.target.value })}
                placeholder="Commentaire de situation..."
                style={{ width: '90%', padding: '6px 10px', borderRadius: '6px', border: '1px solid var(--main-border)', background: 'var(--main-surface)', color: 'var(--main-text)', fontSize: '0.78rem', textAlign: 'center' }}
              />
            </div>
          </div>
        </div>

        {/* 2. SECTION ACTIONS, RÉALISATIONS ET DÉCISIONS */}
        <div style={{ background: 'var(--main-surface)', borderRadius: '20px', border: '1px solid var(--main-border)', overflow: 'hidden', boxShadow: 'var(--card-shadow)' }}>
          <div style={{ background: 'var(--excel-color)', padding: '16px 24px', display: 'flex', justifyContent: 'space-between', alignItems: 'center', color: '#fff' }}>
            <span style={{ fontWeight: 900, fontSize: '0.9rem', textTransform: 'uppercase', letterSpacing: '0.05em' }}>
              Actions, Réalisations, Points d'Attention et Décisions
            </span>
            <button
              onClick={() => {
                const nextId = actions.length + 1;
                setActions([
                  ...actions,
                  { id: nextId, action: 'Nouvelle action stratégique', actor: 'Responsable', deadline: '31/12/2022', status: 'En cours', weeklyUpdate: 'En préparation' }
                ]);
              }}
              style={{
                padding: '6px 14px',
                background: 'rgba(255, 255, 255, 0.2)',
                color: '#fff',
                border: '1px solid rgba(255, 255, 255, 0.4)',
                borderRadius: '8px',
                fontWeight: 800,
                fontSize: '0.78rem',
                cursor: 'pointer'
              }}
            >
              + Ajouter une action
            </button>
          </div>

          <div style={{ overflowX: 'auto' }}>
            <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '0.88rem' }}>
              <thead>
                <tr style={{ background: 'rgba(0,0,0,0.02)', borderBottom: '2px solid var(--main-border)', color: 'var(--text-muted)' }}>
                  <th style={{ padding: '12px 16px', textAlign: 'center', width: '50px' }}>#</th>
                  <th style={{ padding: '12px 16px', textAlign: 'left' }}>Action / Réalisation</th>
                  <th style={{ padding: '12px 16px', textAlign: 'left', width: '180px' }}>Acteur</th>
                  <th style={{ padding: '12px 16px', textAlign: 'center', width: '130px' }}>À faire pour le</th>
                  <th style={{ padding: '12px 16px', textAlign: 'center', width: '140px' }}>Statut</th>
                  <th style={{ padding: '12px 16px', textAlign: 'left' }}>Mise à jour de la semaine</th>
                  <th style={{ padding: '12px 16px', textAlign: 'center', width: '60px' }}>Action</th>
                </tr>
              </thead>
              <tbody>
                {actions.map((item, idx) => {
                  const badge = getStatusBadge(item.status);
                  return (
                    <tr key={item.id} style={{ borderBottom: '1px solid var(--main-border)' }}>
                      <td style={{ padding: '10px 16px', textAlign: 'center', fontWeight: 800, color: 'var(--text-muted)' }}>
                        {idx + 1}
                      </td>
                      <td style={{ padding: '10px 16px' }}>
                        <input
                          type="text"
                          value={item.action}
                          onChange={(e) => {
                            const next = [...actions];
                            next[idx].action = e.target.value;
                            setActions(next);
                          }}
                          style={{ width: '100%', background: 'transparent', border: '1px solid transparent', color: 'var(--main-text)', fontWeight: 700, padding: '4px 6px', borderRadius: '4px' }}
                        />
                      </td>
                      <td style={{ padding: '10px 16px' }}>
                        <input
                          type="text"
                          value={item.actor}
                          onChange={(e) => {
                            const next = [...actions];
                            next[idx].actor = e.target.value;
                            setActions(next);
                          }}
                          style={{ width: '100%', background: 'transparent', border: '1px solid transparent', color: 'var(--main-text)', padding: '4px 6px', borderRadius: '4px' }}
                        />
                      </td>
                      <td style={{ padding: '10px 16px', textAlign: 'center' }}>
                        <input
                          type="text"
                          value={item.deadline}
                          onChange={(e) => {
                            const next = [...actions];
                            next[idx].deadline = e.target.value;
                            setActions(next);
                          }}
                          style={{ width: '100px', textAlign: 'center', background: 'transparent', border: '1px solid transparent', color: 'var(--main-text)', fontWeight: 700 }}
                        />
                      </td>
                      <td style={{ padding: '10px 16px', textAlign: 'center' }}>
                        <button
                          onClick={() => {
                            const next = [...actions];
                            next[idx].status = cycleActionStatus(item.status);
                            setActions(next);
                          }}
                          style={{
                            padding: '4px 12px',
                            borderRadius: '20px',
                            border: 'none',
                            background: badge.bg,
                            color: badge.color,
                            fontWeight: 800,
                            fontSize: '0.72rem',
                            cursor: 'pointer'
                          }}
                        >
                          {badge.label}
                        </button>
                      </td>
                      <td style={{ padding: '10px 16px' }}>
                        <input
                          type="text"
                          value={item.weeklyUpdate}
                          onChange={(e) => {
                            const next = [...actions];
                            next[idx].weeklyUpdate = e.target.value;
                            setActions(next);
                          }}
                          style={{ width: '100%', background: 'transparent', border: '1px solid transparent', color: 'var(--text-muted)', fontStyle: 'italic', padding: '4px 6px', borderRadius: '4px' }}
                        />
                      </td>
                      <td style={{ padding: '10px 16px', textAlign: 'center' }}>
                        <button
                          onClick={() => setActions(actions.filter((_, i) => i !== idx))}
                          style={{ background: 'transparent', border: 'none', color: '#ef4444', cursor: 'pointer', fontSize: '1rem' }}
                          title="Supprimer"
                        >
                          ✕
                        </button>
                      </td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
        </div>

        {/* 3. SECTION INFORMATIONS CLÉS & DÉCISIONS ATTENDUES */}
        <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '2rem' }}>
          {/* Informations importantes de la semaine */}
          <div style={{ background: 'var(--main-surface)', borderRadius: '20px', border: '1px solid var(--main-border)', padding: '2rem', boxShadow: 'var(--card-shadow)' }}>
            <h3 style={{ margin: '0 0 1rem', fontSize: '1rem', fontWeight: 900, color: '#3b82f6', textTransform: 'uppercase', display: 'flex', alignItems: 'center', gap: '8px' }}>
              <span>💡</span> Informations Importantes de la Semaine
            </h3>
            <textarea
              rows={6}
              value={importantInfo}
              onChange={(e) => setImportantInfo(e.target.value)}
              placeholder="Notes et avancements de la semaine..."
              style={{
                width: '100%',
                padding: '12px',
                borderRadius: '12px',
                border: '1px solid var(--main-border)',
                background: 'var(--main-bg)',
                color: 'var(--main-text)',
                fontSize: '0.88rem',
                lineHeight: '1.6',
                fontFamily: 'inherit',
                resize: 'vertical'
              }}
            />
          </div>

          {/* Décisions attendues */}
          <div style={{ background: 'var(--main-surface)', borderRadius: '20px', border: '1px solid var(--main-border)', padding: '2rem', boxShadow: 'var(--card-shadow)' }}>
            <h3 style={{ margin: '0 0 1rem', fontSize: '1rem', fontWeight: 900, color: '#f59e0b', textTransform: 'uppercase', display: 'flex', alignItems: 'center', gap: '8px' }}>
              <span>🎯</span> Décisions Attendues (Comité de Pilotage / Sponsor)
            </h3>
            <textarea
              rows={6}
              value={expectedDecisions}
              onChange={(e) => setExpectedDecisions(e.target.value)}
              placeholder="Décisions requises pour débloquer les prochaines étapes..."
              style={{
                width: '100%',
                padding: '12px',
                borderRadius: '12px',
                border: '1px solid var(--main-border)',
                background: 'var(--main-bg)',
                color: 'var(--main-text)',
                fontSize: '0.88rem',
                lineHeight: '1.6',
                fontFamily: 'inherit',
                resize: 'vertical'
              }}
            />
          </div>
        </div>

        {/* 4. SECTION RISQUES ET POINTS D'ATTENTION */}
        <div style={{ background: 'var(--main-surface)', borderRadius: '20px', border: '1px solid var(--main-border)', overflow: 'hidden', boxShadow: 'var(--card-shadow)' }}>
          <div style={{ background: 'linear-gradient(135deg, #ef4444, #dc2626)', padding: '16px 24px', display: 'flex', justifyContent: 'space-between', alignItems: 'center', color: '#fff' }}>
            <span style={{ fontWeight: 900, fontSize: '0.9rem', textTransform: 'uppercase', letterSpacing: '0.05em' }}>
              ⚠️ Risques Principaux et Points d’Attention
            </span>
            <button
              onClick={() => {
                const nextId = risks.length + 1;
                setRisks([
                  ...risks,
                  { id: nextId, risk: "Nouveau risque identifié", actionProposal: "Plan de remédiation", actor: "Responsable" }
                ]);
              }}
              style={{
                padding: '6px 14px',
                background: 'rgba(255, 255, 255, 0.2)',
                color: '#fff',
                border: '1px solid rgba(255, 255, 255, 0.4)',
                borderRadius: '8px',
                fontWeight: 800,
                fontSize: '0.78rem',
                cursor: 'pointer'
              }}
            >
              + Ajouter un risque
            </button>
          </div>

          <div style={{ overflowX: 'auto' }}>
            <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '0.88rem' }}>
              <thead>
                <tr style={{ background: 'rgba(0,0,0,0.02)', borderBottom: '2px solid var(--main-border)', color: 'var(--text-muted)' }}>
                  <th style={{ padding: '12px 16px', textAlign: 'left', width: '30%' }}>Risques et Points d’Attention</th>
                  <th style={{ padding: '12px 16px', textAlign: 'left', width: '45%' }}>Proposition d'Actions</th>
                  <th style={{ padding: '12px 16px', textAlign: 'left', width: '20%' }}>Acteur</th>
                  <th style={{ padding: '12px 16px', textAlign: 'center', width: '5%' }}>Action</th>
                </tr>
              </thead>
              <tbody>
                {risks.map((item, idx) => (
                  <tr key={item.id} style={{ borderBottom: '1px solid var(--main-border)' }}>
                    <td style={{ padding: '10px 16px' }}>
                      <input
                        type="text"
                        value={item.risk}
                        onChange={(e) => {
                          const next = [...risks];
                          next[idx].risk = e.target.value;
                          setRisks(next);
                        }}
                        style={{ width: '100%', background: 'transparent', border: '1px solid transparent', color: 'var(--main-text)', fontWeight: 700, padding: '4px 6px', borderRadius: '4px' }}
                      />
                    </td>
                    <td style={{ padding: '10px 16px' }}>
                      <input
                        type="text"
                        value={item.actionProposal}
                        onChange={(e) => {
                          const next = [...risks];
                          next[idx].actionProposal = e.target.value;
                          setRisks(next);
                        }}
                        style={{ width: '100%', background: 'transparent', border: '1px solid transparent', color: 'var(--main-text)', padding: '4px 6px', borderRadius: '4px' }}
                      />
                    </td>
                    <td style={{ padding: '10px 16px' }}>
                      <input
                        type="text"
                        value={item.actor}
                        onChange={(e) => {
                          const next = [...risks];
                          next[idx].actor = e.target.value;
                          setRisks(next);
                        }}
                        style={{ width: '100%', background: 'transparent', border: '1px solid transparent', color: 'var(--text-muted)', fontWeight: 600, padding: '4px 6px', borderRadius: '4px' }}
                      />
                    </td>
                    <td style={{ padding: '10px 16px', textAlign: 'center' }}>
                      <button
                        onClick={() => setRisks(risks.filter((_, i) => i !== idx))}
                        style={{ background: 'transparent', border: 'none', color: '#ef4444', cursor: 'pointer', fontSize: '1rem' }}
                        title="Supprimer"
                      >
                        ✕
                      </button>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>
      </div>
    </div>
  );
};
