import React, { useState, useEffect, useMemo } from 'react';
import {
  ResponsiveContainer,
  PieChart,
  Pie,
  Cell,
  LineChart,
  Line,
  BarChart,
  Bar,
  XAxis,
  YAxis,
  Tooltip,
  Legend,
  CartesianGrid
} from 'recharts';

interface TableauDeBordMatrixProps {
  projectId: string;
  fileId: string;
  currentDossierId: string | null;
  activeSubTab?: string;
}

interface MonthlyDeviation {
  month: string;
  scheduleDev: number; // % écart de délai
  budgetDev: number;   // % écart de budget
}

interface MonthlyCost {
  month: string;
  cost: number;
}

interface PhaseProgress {
  phase: string;
  progress: number;
}

const DEFAULT_MONTHLY_DEV: MonthlyDeviation[] = [
  { month: 'mars-22', scheduleDev: 0.2, budgetDev: 2.0 },
  { month: 'avr.-22', scheduleDev: 5.0, budgetDev: 5.0 },
  { month: 'mai-22', scheduleDev: 8.0, budgetDev: 9.0 },
  { month: 'juin-22', scheduleDev: 12.2, budgetDev: 12.3 },
  { month: 'juil.-22', scheduleDev: 16.1, budgetDev: -15.0 },
  { month: 'août-22', scheduleDev: 20.1, budgetDev: 19.3 },
  { month: 'sept.-22', scheduleDev: 24.0, budgetDev: 22.8 },
  { month: 'oct.-22', scheduleDev: -40.0, budgetDev: 10.0 },
  { month: 'nov.-22', scheduleDev: 20.1, budgetDev: 19.3 },
  { month: 'déc.-22', scheduleDev: 35.7, budgetDev: 50.0 },
  { month: 'janv.-23', scheduleDev: 39.7, budgetDev: 10.0 },
  { month: 'févr.-23', scheduleDev: 43.6, budgetDev: 40.3 },
  { month: 'mars-23', scheduleDev: 47.5, budgetDev: 43.8 },
  { month: 'avr.-23', scheduleDev: 51.4, budgetDev: 47.3 },
  { month: 'mai-23', scheduleDev: 50.0, budgetDev: 60.0 }
];

const DEFAULT_MONTHLY_COSTS: MonthlyCost[] = [
  { month: 'mars-22', cost: 82000 },
  { month: 'avr.-22', cost: 82000 },
  { month: 'mai-22', cost: 82000 },
  { month: 'juin-22', cost: 82000 },
  { month: 'juil.-22', cost: 82000 },
  { month: 'août-22', cost: 20000 },
  { month: 'sept.-22', cost: 90000 },
  { month: 'oct.-22', cost: 90000 },
  { month: 'nov.-22', cost: 25000 },
  { month: 'déc.-22', cost: 15000 },
  { month: 'janv.-23', cost: 20000 },
  { month: 'févr.-23', cost: 20115 }
];

const TASK_STATUS_COLORS: { [key: string]: string } = {
  'À faire (TODO)': '#94a3b8',
  'En cours': '#3b82f6',
  'En attente (Hold)': '#f59e0b',
  'Terminé (Done)': '#10b981'
};

export const TableauDeBordMatrix: React.FC<TableauDeBordMatrixProps> = ({
  projectId,
  fileId,
  currentDossierId,
  activeSubTab = 'global'
}) => {
  const dossierSubKey = currentDossierId ? `-${currentDossierId}` : '';
  const storageKey = `faab-tb-dashboard-${projectId}-${fileId}${dossierSubKey}`;

  // State
  const [activeTab, setActiveTab] = useState<string>(activeSubTab || 'global');

  // Budget & Schedule settings
  const [initialBudget, setInitialBudget] = useState<number>(1000000);
  const [startDateInit, setStartDateInit] = useState<string>('2022-03-15');
  const [endDateInit, setEndDateInit] = useState<string>('2022-11-29');
  const [endDateEst, setEndDateEst] = useState<string>('2023-01-15');

  // Tasks counts
  const [taskStats, setTaskStats] = useState({
    todo: 19,
    inProgress: 9,
    onHold: 5,
    done: 2
  });

  // Phases progress
  const [phases, setPhases] = useState<PhaseProgress[]>([
    { phase: '1. Initialisation', progress: 76 },
    { phase: '2. Planification', progress: 60 },
    { phase: '3. Exécution', progress: 31 },
    { phase: '4. Contrôle & Suivi', progress: 32 },
    { phase: '5. Clôture', progress: 36 }
  ]);

  // Monthly costs & deviations
  const [monthlyCosts, setMonthlyCosts] = useState<MonthlyCost[]>(() => {
    const saved = localStorage.getItem(`${storageKey}-costs`);
    if (saved) {
      try { return JSON.parse(saved); } catch (e) { }
    }
    return DEFAULT_MONTHLY_COSTS;
  });

  const [monthlyDev, setMonthlyDev] = useState<MonthlyDeviation[]>(() => {
    const saved = localStorage.getItem(`${storageKey}-dev`);
    if (saved) {
      try { return JSON.parse(saved); } catch (e) { }
    }
    return DEFAULT_MONTHLY_DEV;
  });

  useEffect(() => {
    if (activeSubTab) setActiveTab(activeSubTab);
  }, [activeSubTab]);

  useEffect(() => {
    localStorage.setItem(`${storageKey}-costs`, JSON.stringify(monthlyCosts));
  }, [monthlyCosts, storageKey]);

  useEffect(() => {
    localStorage.setItem(`${storageKey}-dev`, JSON.stringify(monthlyDev));
  }, [monthlyDev, storageKey]);

  // Automatics calculations
  const totalTasks = taskStats.todo + taskStats.inProgress + taskStats.onHold + taskStats.done;
  const unfinishedTasks = taskStats.todo + taskStats.inProgress + taskStats.onHold;
  const progressRate = totalTasks > 0 ? Math.round((taskStats.done / totalTasks) * 100) : 47;

  const totalCostConsumed = useMemo(() => {
    return monthlyCosts.reduce((acc, curr) => acc + curr.cost, 0);
  }, [monthlyCosts]);

  // Re-estimated total budget (EAC = AC / progress)
  const reestimatedBudget = useMemo(() => {
    const pct = progressRate > 0 ? progressRate / 100 : 0.47;
    return Math.round(totalCostConsumed / pct);
  }, [totalCostConsumed, progressRate]);

  // Budget deviation %
  const budgetDeviationPct = useMemo(() => {
    if (initialBudget <= 0) return 0;
    return Math.round(((reestimatedBudget - initialBudget) / initialBudget) * 100);
  }, [reestimatedBudget, initialBudget]);

  // Schedule deviation % (difference between dates in days)
  const scheduleDeviationPct = useMemo(() => {
    const d1 = new Date(startDateInit).getTime();
    const d2 = new Date(endDateInit).getTime();
    const d3 = new Date(endDateEst).getTime();
    const initDuration = Math.max(1, Math.round((d2 - d1) / (1000 * 3600 * 24)));
    const estDuration = Math.max(1, Math.round((d3 - d1) / (1000 * 3600 * 24)));
    return Math.round(((estDuration - initDuration) / initDuration) * 100);
  }, [startDateInit, endDateInit, endDateEst]);

  // Pie chart data
  const pieData = [
    { name: 'À faire (TODO)', value: taskStats.todo },
    { name: 'En cours', value: taskStats.inProgress },
    { name: 'En attente (Hold)', value: taskStats.onHold },
    { name: 'Terminé (Done)', value: taskStats.done }
  ];

  return (
    <div style={{ flex: 1, display: 'flex', flexDirection: 'column', background: 'var(--main-bg)', color: 'var(--main-text)', overflowY: 'auto' }}>
      {/* Top Header */}
      <div style={{ padding: '1.25rem 2rem', background: 'var(--main-surface)', borderBottom: '1px solid var(--main-border)', display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
        <div>
          <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
            <span style={{ fontSize: '1.3rem' }}>📉</span>
            <h1 style={{ fontSize: '1.25rem', fontWeight: 900, margin: 0, color: 'var(--main-text)' }}>
              29 — Tableau de Bord Global & Pilotage
            </h1>
            <span style={{ padding: '4px 10px', background: 'rgba(59, 130, 246, 0.12)', color: '#3b82f6', borderRadius: '20px', fontSize: '0.72rem', fontWeight: 800 }}>
              EXCEL 29 DYNAMIQUE
            </span>
          </div>
          <p style={{ margin: '4px 0 0', fontSize: '0.8rem', color: 'var(--text-muted)' }}>
            Kit Chef de Projet • Suivi consolidé des délais, coûts, activités et écarts avec calculs temps réel
          </p>
        </div>
      </div>

      {/* Sub-tabs bar */}
      <div style={{ display: 'flex', gap: '8px', padding: '0.75rem 2rem', background: 'var(--main-surface)', borderBottom: '1px solid var(--main-border)', overflowX: 'auto' }}>
        {[
          { id: 'global', label: '📊 Dashboard Consolidé' },
          { id: 'deviations', label: '📈 Écarts Délais & Coûts' },
          { id: 'costs', label: '💶 Coûts Consommés Mensuels' },
          { id: 'tasks', label: '📋 Pilotage des Tâches & Phases' }
        ].map(tab => (
          <button
            key={tab.id}
            onClick={() => setActiveTab(tab.id)}
            style={{
              padding: '8px 16px',
              borderRadius: '8px',
              border: activeTab === tab.id ? '1px solid #10b981' : '1px solid var(--main-border)',
              background: activeTab === tab.id ? 'rgba(16, 185, 129, 0.12)' : 'var(--main-bg)',
              color: activeTab === tab.id ? '#10b981' : 'var(--main-text)',
              fontWeight: 800,
              fontSize: '0.8rem',
              cursor: 'pointer',
              whiteSpace: 'nowrap',
              transition: 'all 0.15s ease'
            }}
          >
            {tab.label}
          </button>
        ))}
      </div>

      {/* Main Content Area */}
      <div style={{ padding: '2rem', display: 'flex', flexDirection: 'column', gap: '2rem' }}>
        {/* TOP KPI CARDS */}
        <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(240px, 1fr))', gap: '1.25rem' }}>
          {/* Card 1: Fin du projet */}
          <div style={{ background: 'var(--main-surface)', padding: '1.5rem', borderRadius: '16px', border: '1px solid var(--main-border)', borderLeft: '5px solid #3b82f6', boxShadow: 'var(--card-shadow)' }}>
            <div style={{ fontSize: '0.75rem', fontWeight: 800, color: 'var(--text-muted)', textTransform: 'uppercase' }}>Estimation Fin Projet</div>
            <div style={{ fontSize: '1.4rem', fontWeight: 950, marginTop: '6px', color: scheduleDeviationPct > 0 ? '#ef4444' : '#10b981' }}>
              {endDateEst}
            </div>
            <div style={{ fontSize: '0.8rem', color: 'var(--text-muted)', marginTop: '4px' }}>
              Prévu initial : <b style={{ color: 'var(--main-text)' }}>{endDateInit}</b> ({scheduleDeviationPct > 0 ? `+${scheduleDeviationPct}% retard` : 'Dans les délais'})
            </div>
          </div>

          {/* Card 2: Taux d'avancement */}
          <div style={{ background: 'var(--main-surface)', padding: '1.5rem', borderRadius: '16px', border: '1px solid var(--main-border)', borderLeft: '5px solid #10b981', boxShadow: 'var(--card-shadow)', display: 'flex', alignItems: 'center', justifyContent: 'space-between' }}>
            <div>
              <div style={{ fontSize: '0.75rem', fontWeight: 800, color: 'var(--text-muted)', textTransform: 'uppercase' }}>Avancement Global</div>
              <div style={{ fontSize: '2.2rem', fontWeight: 950, marginTop: '4px', color: '#10b981' }}>
                {progressRate}%
              </div>
              <div style={{ fontSize: '0.8rem', color: 'var(--text-muted)' }}>
                {taskStats.done} terminées / {totalTasks} tâches
              </div>
            </div>
            <div style={{ width: '60px', height: '60px', borderRadius: '50%', background: `conic-gradient(#10b981 ${progressRate}%, rgba(16,185,129,0.15) 0)`, display: 'flex', alignItems: 'center', justifyContent: 'center' }}>
              <div style={{ width: '46px', height: '46px', borderRadius: '50%', background: 'var(--main-surface)', display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: '0.7rem', fontWeight: 900 }}>
                {progressRate}%
              </div>
            </div>
          </div>

          {/* Card 3: Budget Total Initial vs Réestimé */}
          <div style={{ background: 'var(--main-surface)', padding: '1.5rem', borderRadius: '16px', border: '1px solid var(--main-border)', borderLeft: '5px solid #f59e0b', boxShadow: 'var(--card-shadow)' }}>
            <div style={{ fontSize: '0.75rem', fontWeight: 800, color: 'var(--text-muted)', textTransform: 'uppercase' }}>Budget Total Estimé (EAC)</div>
            <div style={{ fontSize: '1.4rem', fontWeight: 950, marginTop: '6px', color: budgetDeviationPct > 0 ? '#f59e0b' : '#10b981' }}>
              {reestimatedBudget.toLocaleString('fr-FR')} €
            </div>
            <div style={{ fontSize: '0.8rem', color: 'var(--text-muted)', marginTop: '4px' }}>
              Initial : <b style={{ color: 'var(--main-text)' }}>{initialBudget.toLocaleString('fr-FR')} €</b> ({budgetDeviationPct > 0 ? `+${budgetDeviationPct}% dérive` : 'Sous contrôle'})
            </div>
          </div>

          {/* Card 4: Budget Consommé */}
          <div style={{ background: 'var(--main-surface)', padding: '1.5rem', borderRadius: '16px', border: '1px solid var(--main-border)', borderLeft: '5px solid #8b5cf6', boxShadow: 'var(--card-shadow)' }}>
            <div style={{ fontSize: '0.75rem', fontWeight: 800, color: 'var(--text-muted)', textTransform: 'uppercase' }}>Budget Consommé à ce jour</div>
            <div style={{ fontSize: '1.4rem', fontWeight: 950, marginTop: '6px', color: '#8b5cf6' }}>
              {totalCostConsumed.toLocaleString('fr-FR')} €
            </div>
            <div style={{ fontSize: '0.8rem', color: 'var(--text-muted)', marginTop: '4px' }}>
              Consommation : <b style={{ color: 'var(--main-text)' }}>{Math.round((totalCostConsumed / initialBudget) * 100)}%</b> du budget initial
            </div>
          </div>
        </div>

        {/* TAB 1: DASHBOARD CONSOLIDÉ */}
        {activeTab === 'global' && (
          <>
            <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '2rem' }}>
              {/* SANTE DU PROJET */}
              <div style={{ background: 'var(--main-surface)', padding: '2rem', borderRadius: '20px', border: '1px solid var(--main-border)', boxShadow: 'var(--card-shadow)' }}>
                <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '1.5rem' }}>
                  <h3 style={{ margin: 0, fontSize: '1rem', fontWeight: 900, color: '#10b981', textTransform: 'uppercase', display: 'flex', alignItems: 'center', gap: '8px' }}>
                    <span>🩺</span> Santé Globale du Projet (Health Card)
                  </h3>
                  <span style={{ fontSize: '0.75rem', padding: '4px 10px', background: 'rgba(16, 185, 129, 0.1)', color: '#10b981', borderRadius: '20px', fontWeight: 800 }}>
                    Calcul automatique
                  </span>
                </div>
                <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '0.88rem' }}>
                  <tbody>
                    <tr style={{ borderBottom: '1px solid var(--main-border)' }}>
                      <td style={{ padding: '14px 0', color: 'var(--text-muted)', fontWeight: 600 }}>Dérive Calendaire</td>
                      <td style={{ padding: '14px 0', textAlign: 'right', fontWeight: 900, color: scheduleDeviationPct > 0 ? '#ef4444' : '#10b981' }}>
                        {scheduleDeviationPct > 0 ? `+${scheduleDeviationPct}% de retard` : 'Conforme au planning'}
                      </td>
                    </tr>
                    <tr style={{ borderBottom: '1px solid var(--main-border)' }}>
                      <td style={{ padding: '14px 0', color: 'var(--text-muted)', fontWeight: 600 }}>Activités en cours & en attente</td>
                      <td style={{ padding: '14px 0', textAlign: 'right', fontWeight: 900, color: '#f59e0b' }}>
                        {unfinishedTasks} activités non terminées
                      </td>
                    </tr>
                    <tr style={{ borderBottom: '1px solid var(--main-border)' }}>
                      <td style={{ padding: '14px 0', color: 'var(--text-muted)', fontWeight: 600 }}>Activités entièrement achevées</td>
                      <td style={{ padding: '14px 0', textAlign: 'right', fontWeight: 900, color: '#10b981' }}>
                        {taskStats.done} activités terminées ({progressRate}%)
                      </td>
                    </tr>
                    <tr>
                      <td style={{ padding: '14px 0', color: 'var(--text-muted)', fontWeight: 600 }}>Dérive Budgétaire (Surconsommation)</td>
                      <td style={{ padding: '14px 0', textAlign: 'right', fontWeight: 900, color: budgetDeviationPct > 0 ? '#ef4444' : '#10b981' }}>
                        {budgetDeviationPct > 0 ? `+${budgetDeviationPct}% de surconsommation` : 'Sous le budget prévisionnel'}
                      </td>
                    </tr>
                  </tbody>
                </table>
              </div>

              {/* PROGRESSION PAR PHASE */}
              <div style={{ background: 'var(--main-surface)', padding: '2rem', borderRadius: '20px', border: '1px solid var(--main-border)', boxShadow: 'var(--card-shadow)' }}>
                <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '1.5rem' }}>
                  <h3 style={{ margin: 0, fontSize: '1rem', fontWeight: 900, color: '#3b82f6', textTransform: 'uppercase', display: 'flex', alignItems: 'center', gap: '8px' }}>
                    <span>🏗️</span> Progression par Phase du Projet
                  </h3>
                  <span style={{ fontSize: '0.75rem', color: 'var(--text-muted)', fontWeight: 700 }}>5 Phases</span>
                </div>
                <div style={{ display: 'flex', flexDirection: 'column', gap: '1.25rem' }}>
                  {phases.map((p, idx) => (
                    <div key={p.phase}>
                      <div style={{ display: 'flex', justifyContent: 'space-between', fontSize: '0.82rem', marginBottom: '6px' }}>
                        <span style={{ fontWeight: 700, color: 'var(--main-text)' }}>{p.phase}</span>
                        <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
                          <input
                            type="number"
                            min="0"
                            max="100"
                            value={p.progress}
                            onChange={(e) => {
                              const val = Math.min(100, Math.max(0, parseInt(e.target.value) || 0));
                              const next = [...phases];
                              next[idx].progress = val;
                              setPhases(next);
                            }}
                            style={{
                              width: '55px',
                              padding: '2px 6px',
                              borderRadius: '6px',
                              border: '1px solid var(--main-border)',
                              background: 'var(--main-bg)',
                              color: 'var(--main-text)',
                              fontSize: '0.78rem',
                              fontWeight: 800,
                              textAlign: 'right'
                            }}
                          />
                          <span style={{ fontWeight: 900, color: p.progress > 70 ? '#10b981' : p.progress > 40 ? '#3b82f6' : '#f59e0b' }}>
                            %
                          </span>
                        </div>
                      </div>
                      <div style={{ height: '8px', background: 'var(--main-bg)', borderRadius: '4px', overflow: 'hidden', border: '1px solid var(--main-border)' }}>
                        <div
                          style={{
                            width: `${p.progress}%`,
                            height: '100%',
                            background: p.progress > 70 ? '#10b981' : p.progress > 40 ? '#3b82f6' : '#f59e0b',
                            borderRadius: '4px',
                            transition: 'width 0.3s ease'
                          }}
                        />
                      </div>
                    </div>
                  ))}
                </div>
              </div>
            </div>

            {/* GRAPHIQUES */}
            <div style={{ display: 'grid', gridTemplateColumns: '1fr 1.3fr', gap: '2rem' }}>
              {/* REPARTITION DES TACHES PIE CHART */}
              <div style={{ background: 'var(--main-surface)', padding: '2rem', borderRadius: '20px', border: '1px solid var(--main-border)', boxShadow: 'var(--card-shadow)' }}>
                <h3 style={{ margin: '0 0 1rem', fontSize: '0.95rem', fontWeight: 900, color: 'var(--main-text)', textTransform: 'uppercase' }}>
                  🎯 Répartition des Activités par Statut
                </h3>
                <div style={{ height: '260px' }}>
                  <ResponsiveContainer width="100%" height="100%">
                    <PieChart>
                      <Pie
                        data={pieData}
                        cx="50%"
                        cy="50%"
                        innerRadius={55}
                        outerRadius={90}
                        paddingAngle={4}
                        dataKey="value"
                      >
                        {pieData.map((entry) => (
                          <Cell key={entry.name} fill={TASK_STATUS_COLORS[entry.name]} />
                        ))}
                      </Pie>
                      <Tooltip />
                      <Legend verticalAlign="bottom" height={36} />
                    </PieChart>
                  </ResponsiveContainer>
                </div>
              </div>

              {/* EVOLUTION DES ECARTS LINE CHART */}
              <div style={{ background: 'var(--main-surface)', padding: '2rem', borderRadius: '20px', border: '1px solid var(--main-border)', boxShadow: 'var(--card-shadow)' }}>
                <h3 style={{ margin: '0 0 1rem', fontSize: '0.95rem', fontWeight: 900, color: 'var(--main-text)', textTransform: 'uppercase' }}>
                  📈 Courbe d'Évolution des Écarts (%)
                </h3>
                <div style={{ height: '260px' }}>
                  <ResponsiveContainer width="100%" height="100%">
                    <LineChart data={monthlyDev} margin={{ top: 10, right: 20, left: 0, bottom: 5 }}>
                      <CartesianGrid strokeDasharray="3 3" stroke="var(--main-border)" opacity={0.5} />
                      <XAxis dataKey="month" stroke="var(--text-muted)" fontSize={11} />
                      <YAxis stroke="var(--text-muted)" fontSize={11} unit="%" />
                      <Tooltip formatter={(value: any) => [`${value}%`]} />
                      <Legend />
                      <Line type="monotone" dataKey="scheduleDev" name="Écart de délai (%)" stroke="#ef4444" strokeWidth={2} dot={{ r: 3 }} />
                      <Line type="monotone" dataKey="budgetDev" name="Écart de budget (%)" stroke="#3b82f6" strokeWidth={2} dot={{ r: 3 }} />
                    </LineChart>
                  </ResponsiveContainer>
                </div>
              </div>
            </div>
          </>
        )}

        {/* TAB 2: HISTORIQUE DES ÉCARTS (DEVIATIONS) */}
        {activeTab === 'deviations' && (
          <div style={{ background: 'var(--main-surface)', borderRadius: '20px', border: '1px solid var(--main-border)', padding: '2rem', boxShadow: 'var(--card-shadow)' }}>
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '1.5rem' }}>
              <div>
                <h3 style={{ margin: 0, fontSize: '1.1rem', fontWeight: 900, color: 'var(--main-text)' }}>
                  Hist. Dev. Time Budget • Matrice des Écarts Mensuels
                </h3>
                <p style={{ margin: '4px 0 0', fontSize: '0.8rem', color: 'var(--text-muted)' }}>
                  Modifiez directement les pourcentages d'écart de délai et de budget pour ajuster les courbes automatiquement
                </p>
              </div>
              <button
                onClick={() => {
                  const nextMonth = `Mois-${monthlyDev.length + 1}`;
                  setMonthlyDev([...monthlyDev, { month: nextMonth, scheduleDev: 0, budgetDev: 0 }]);
                }}
                style={{
                  padding: '8px 16px',
                  background: '#10b981',
                  color: '#fff',
                  border: 'none',
                  borderRadius: '8px',
                  fontWeight: 800,
                  fontSize: '0.8rem',
                  cursor: 'pointer'
                }}
              >
                + Ajouter un mois
              </button>
            </div>

            <div style={{ overflowX: 'auto' }}>
              <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '0.88rem' }}>
                <thead>
                  <tr style={{ background: 'rgba(0,0,0,0.03)', borderBottom: '2px solid var(--main-border)', color: 'var(--text-muted)' }}>
                    <th style={{ padding: '12px 16px', textAlign: 'left', fontWeight: 800 }}>Période / Mois</th>
                    <th style={{ padding: '12px 16px', textAlign: 'center', fontWeight: 800 }}>Écart de délai (%)</th>
                    <th style={{ padding: '12px 16px', textAlign: 'center', fontWeight: 800 }}>Écart de budget (%)</th>
                    <th style={{ padding: '12px 16px', textAlign: 'center', fontWeight: 800 }}>Statut Délais</th>
                    <th style={{ padding: '12px 16px', textAlign: 'center', fontWeight: 800 }}>Statut Budget</th>
                  </tr>
                </thead>
                <tbody>
                  {monthlyDev.map((item, idx) => (
                    <tr key={idx} style={{ borderBottom: '1px solid var(--main-border)' }}>
                      <td style={{ padding: '10px 16px', fontWeight: 700 }}>
                        <input
                          type="text"
                          value={item.month}
                          onChange={(e) => {
                            const next = [...monthlyDev];
                            next[idx].month = e.target.value;
                            setMonthlyDev(next);
                          }}
                          style={{
                            background: 'transparent',
                            border: '1px solid transparent',
                            color: 'var(--main-text)',
                            fontWeight: 700,
                            padding: '4px 8px',
                            borderRadius: '4px'
                          }}
                        />
                      </td>
                      <td style={{ padding: '10px 16px', textAlign: 'center' }}>
                        <input
                          type="number"
                          step="0.1"
                          value={item.scheduleDev}
                          onChange={(e) => {
                            const next = [...monthlyDev];
                            next[idx].scheduleDev = parseFloat(e.target.value) || 0;
                            setMonthlyDev(next);
                          }}
                          style={{
                            width: '90px',
                            padding: '6px 10px',
                            borderRadius: '6px',
                            border: '1px solid var(--main-border)',
                            background: 'var(--main-bg)',
                            color: 'var(--main-text)',
                            textAlign: 'right',
                            fontWeight: 800
                          }}
                        />
                        <span style={{ marginLeft: '4px', fontSize: '0.8rem', color: 'var(--text-muted)' }}>%</span>
                      </td>
                      <td style={{ padding: '10px 16px', textAlign: 'center' }}>
                        <input
                          type="number"
                          step="0.1"
                          value={item.budgetDev}
                          onChange={(e) => {
                            const next = [...monthlyDev];
                            next[idx].budgetDev = parseFloat(e.target.value) || 0;
                            setMonthlyDev(next);
                          }}
                          style={{
                            width: '90px',
                            padding: '6px 10px',
                            borderRadius: '6px',
                            border: '1px solid var(--main-border)',
                            background: 'var(--main-bg)',
                            color: 'var(--main-text)',
                            textAlign: 'right',
                            fontWeight: 800
                          }}
                        />
                        <span style={{ marginLeft: '4px', fontSize: '0.8rem', color: 'var(--text-muted)' }}>%</span>
                      </td>
                      <td style={{ padding: '10px 16px', textAlign: 'center' }}>
                        <span style={{
                          padding: '4px 10px',
                          borderRadius: '20px',
                          fontSize: '0.72rem',
                          fontWeight: 800,
                          background: item.scheduleDev > 0 ? 'rgba(239, 68, 68, 0.12)' : 'rgba(16, 185, 129, 0.12)',
                          color: item.scheduleDev > 0 ? '#ef4444' : '#10b981'
                        }}>
                          {item.scheduleDev > 0 ? `+${item.scheduleDev}% Retard` : `${item.scheduleDev}% Avance`}
                        </span>
                      </td>
                      <td style={{ padding: '10px 16px', textAlign: 'center' }}>
                        <span style={{
                          padding: '4px 10px',
                          borderRadius: '20px',
                          fontSize: '0.72rem',
                          fontWeight: 800,
                          background: item.budgetDev > 0 ? 'rgba(245, 158, 11, 0.12)' : 'rgba(16, 185, 129, 0.12)',
                          color: item.budgetDev > 0 ? '#f59e0b' : '#10b981'
                        }}>
                          {item.budgetDev > 0 ? `+${item.budgetDev}% Dépassement` : `${item.budgetDev}% Économie`}
                        </span>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>
        )}

        {/* TAB 3: COÛTS CONSOMMÉS (COST TRACKING) */}
        {activeTab === 'costs' && (
          <div style={{ background: 'var(--main-surface)', borderRadius: '20px', border: '1px solid var(--main-border)', padding: '2rem', boxShadow: 'var(--card-shadow)' }}>
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '1.5rem' }}>
              <div>
                <h3 style={{ margin: 0, fontSize: '1.1rem', fontWeight: 900, color: 'var(--main-text)' }}>
                  Cost Tracking • Coûts Consommés par Mois
                </h3>
                <p style={{ margin: '4px 0 0', fontSize: '0.8rem', color: 'var(--text-muted)' }}>
                  Total Consommé à ce jour : <b style={{ color: '#10b981' }}>{totalCostConsumed.toLocaleString('fr-FR')} €</b>
                </p>
              </div>
              <button
                onClick={() => {
                  const nextMonth = `Mois-${monthlyCosts.length + 1}`;
                  setMonthlyCosts([...monthlyCosts, { month: nextMonth, cost: 0 }]);
                }}
                style={{
                  padding: '8px 16px',
                  background: '#10b981',
                  color: '#fff',
                  border: 'none',
                  borderRadius: '8px',
                  fontWeight: 800,
                  fontSize: '0.8rem',
                  cursor: 'pointer'
                }}
              >
                + Ajouter une ligne de coût
              </button>
            </div>

            <div style={{ height: '240px', marginBottom: '2rem' }}>
              <ResponsiveContainer width="100%" height="100%">
                <BarChart data={monthlyCosts}>
                  <CartesianGrid strokeDasharray="3 3" stroke="var(--main-border)" opacity={0.5} />
                  <XAxis dataKey="month" stroke="var(--text-muted)" fontSize={11} />
                  <YAxis stroke="var(--text-muted)" fontSize={11} tickFormatter={(v) => `${v / 1000}k`} />
                  <Tooltip formatter={(value: any) => [`${Number(value).toLocaleString('fr-FR')} €`, 'Coût Consommé']} />
                  <Bar dataKey="cost" fill="#8b5cf6" radius={[4, 4, 0, 0]} />
                </BarChart>
              </ResponsiveContainer>
            </div>

            <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '0.88rem' }}>
              <thead>
                <tr style={{ background: 'rgba(0,0,0,0.03)', borderBottom: '2px solid var(--main-border)', color: 'var(--text-muted)' }}>
                  <th style={{ padding: '12px 16px', textAlign: 'left', fontWeight: 800 }}>Mois / Échéance</th>
                  <th style={{ padding: '12px 16px', textAlign: 'right', fontWeight: 800 }}>Coûts consommés (€)</th>
                  <th style={{ padding: '12px 16px', textAlign: 'right', fontWeight: 800 }}>Cumul Consommé (€)</th>
                </tr>
              </thead>
              <tbody>
                {monthlyCosts.map((item, idx) => {
                  const runningTotal = monthlyCosts.slice(0, idx + 1).reduce((acc, c) => acc + c.cost, 0);
                  return (
                    <tr key={idx} style={{ borderBottom: '1px solid var(--main-border)' }}>
                      <td style={{ padding: '10px 16px', fontWeight: 700 }}>
                        <input
                          type="text"
                          value={item.month}
                          onChange={(e) => {
                            const next = [...monthlyCosts];
                            next[idx].month = e.target.value;
                            setMonthlyCosts(next);
                          }}
                          style={{
                            background: 'transparent',
                            border: '1px solid transparent',
                            color: 'var(--main-text)',
                            fontWeight: 700,
                            padding: '4px 8px',
                            borderRadius: '4px'
                          }}
                        />
                      </td>
                      <td style={{ padding: '10px 16px', textAlign: 'right' }}>
                        <input
                          type="number"
                          value={item.cost}
                          onChange={(e) => {
                            const next = [...monthlyCosts];
                            next[idx].cost = parseFloat(e.target.value) || 0;
                            setMonthlyCosts(next);
                          }}
                          style={{
                            width: '120px',
                            padding: '6px 10px',
                            borderRadius: '6px',
                            border: '1px solid var(--main-border)',
                            background: 'var(--main-bg)',
                            color: 'var(--main-text)',
                            textAlign: 'right',
                            fontWeight: 800
                          }}
                        />
                      </td>
                      <td style={{ padding: '10px 16px', textAlign: 'right', fontWeight: 900, color: 'var(--text-muted)' }}>
                        {runningTotal.toLocaleString('fr-FR')} €
                      </td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
        )}

        {/* TAB 4: PILOTAGE DES TACHES & SETTINGS */}
        {activeTab === 'tasks' && (
          <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '2rem' }}>
            {/* Taches par statut */}
            <div style={{ background: 'var(--main-surface)', borderRadius: '20px', border: '1px solid var(--main-border)', padding: '2rem', boxShadow: 'var(--card-shadow)' }}>
              <h3 style={{ margin: '0 0 1.5rem', fontSize: '1.1rem', fontWeight: 900, color: 'var(--main-text)' }}>
                Comptage des Activités par Statut
              </h3>
              <div style={{ display: 'flex', flexDirection: 'column', gap: '1.25rem' }}>
                {[
                  { key: 'todo', label: 'À faire (TODO)', color: '#94a3b8' },
                  { key: 'inProgress', label: 'En cours', color: '#3b82f6' },
                  { key: 'onHold', label: 'En attente (Hold)', color: '#f59e0b' },
                  { key: 'done', label: 'Terminé (Done)', color: '#10b981' }
                ].map(item => (
                  <div key={item.key} style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                    <div style={{ display: 'flex', alignItems: 'center', gap: '10px' }}>
                      <div style={{ width: '12px', height: '12px', borderRadius: '50%', background: item.color }} />
                      <span style={{ fontWeight: 700, fontSize: '0.9rem' }}>{item.label}</span>
                    </div>
                    <input
                      type="number"
                      min="0"
                      value={(taskStats as any)[item.key]}
                      onChange={(e) => {
                        const val = Math.max(0, parseInt(e.target.value) || 0);
                        setTaskStats({ ...taskStats, [item.key]: val });
                      }}
                      style={{
                        width: '70px',
                        padding: '6px 10px',
                        borderRadius: '6px',
                        border: '1px solid var(--main-border)',
                        background: 'var(--main-bg)',
                        color: 'var(--main-text)',
                        textAlign: 'right',
                        fontWeight: 900
                      }}
                    />
                  </div>
                ))}
              </div>
            </div>

            {/* Parametres projet */}
            <div style={{ background: 'var(--main-surface)', borderRadius: '20px', border: '1px solid var(--main-border)', padding: '2rem', boxShadow: 'var(--card-shadow)' }}>
              <h3 style={{ margin: '0 0 1.5rem', fontSize: '1.1rem', fontWeight: 900, color: 'var(--main-text)' }}>
                Paramètres Généraux du Projet (Settings)
              </h3>
              <div style={{ display: 'flex', flexDirection: 'column', gap: '1.25rem' }}>
                <div>
                  <label style={{ display: 'block', fontSize: '0.78rem', fontWeight: 800, color: 'var(--text-muted)', marginBottom: '4px' }}>
                    Budget Total Initial Prévu (€)
                  </label>
                  <input
                    type="number"
                    value={initialBudget}
                    onChange={(e) => setInitialBudget(parseFloat(e.target.value) || 0)}
                    style={{
                      width: '100%',
                      padding: '8px 12px',
                      borderRadius: '8px',
                      border: '1px solid var(--main-border)',
                      background: 'var(--main-bg)',
                      color: 'var(--main-text)',
                      fontWeight: 800
                    }}
                  />
                </div>
                <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '1rem' }}>
                  <div>
                    <label style={{ display: 'block', fontSize: '0.78rem', fontWeight: 800, color: 'var(--text-muted)', marginBottom: '4px' }}>
                      Date Début Initiale
                    </label>
                    <input
                      type="date"
                      value={startDateInit}
                      onChange={(e) => setStartDateInit(e.target.value)}
                      style={{
                        width: '100%',
                        padding: '8px 12px',
                        borderRadius: '8px',
                        border: '1px solid var(--main-border)',
                        background: 'var(--main-bg)',
                        color: 'var(--main-text)'
                      }}
                    />
                  </div>
                  <div>
                    <label style={{ display: 'block', fontSize: '0.78rem', fontWeight: 800, color: 'var(--text-muted)', marginBottom: '4px' }}>
                      Date Fin Initiale
                    </label>
                    <input
                      type="date"
                      value={endDateInit}
                      onChange={(e) => setEndDateInit(e.target.value)}
                      style={{
                        width: '100%',
                        padding: '8px 12px',
                        borderRadius: '8px',
                        border: '1px solid var(--main-border)',
                        background: 'var(--main-bg)',
                        color: 'var(--main-text)'
                      }}
                    />
                  </div>
                </div>
                <div>
                  <label style={{ display: 'block', fontSize: '0.78rem', fontWeight: 800, color: 'var(--text-muted)', marginBottom: '4px' }}>
                    Date Fin Ré-estimée (À ce jour)
                  </label>
                  <input
                    type="date"
                    value={endDateEst}
                    onChange={(e) => setEndDateEst(e.target.value)}
                    style={{
                      width: '100%',
                      padding: '8px 12px',
                      borderRadius: '8px',
                      border: '1px solid var(--main-border)',
                      background: 'var(--main-bg)',
                      color: 'var(--main-text)'
                    }}
                  />
                </div>
              </div>
            </div>
          </div>
        )}
      </div>
    </div>
  );
};
