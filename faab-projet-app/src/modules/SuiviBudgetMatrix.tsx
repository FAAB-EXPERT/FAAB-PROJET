import React, { useState, useEffect, useMemo } from 'react';
import { ResponsiveContainer, LineChart, Line, XAxis, YAxis, Tooltip, Legend, CartesianGrid } from 'recharts';

interface SuiviBudgetMatrixProps {
  projectId: string;
  fileId: string;
  currentDossierId: string | null;
  activeSubTab?: string;
}

const MONTHS = [
  'mars-22', 'avr.-22', 'mai-22', 'juin-22', 'juil.-22', 'août-22',
  'sept.-22', 'oct.-22', 'nov.-22', 'déc.-22', 'janv.-23', 'févr.-23'
];

interface MonthlyData {
  month: string;
  cr: number; // Coût Réel cumulé
  va: number; // Valeur Acquise cumulée
  vp: number; // Valeur Planifiée cumulée
}

const DEFAULT_MONTHLY: MonthlyData[] = [
  { month: 'mars-22', cr: 4000, va: 3000, vp: 5000 },
  { month: 'avr.-22', cr: 22000, va: 18880, vp: 22000 },
  { month: 'mai-22', cr: 39000, va: 35760, vp: 39000 },
  { month: 'juin-22', cr: 55500, va: 53000, vp: 56000 },
  { month: 'juil.-22', cr: 93500, va: 93660, vp: 94000 },
  { month: 'août-22', cr: 131500, va: 131320, vp: 132000 },
  { month: 'sept.-22', cr: 169500, va: 170000, vp: 170000 },
  { month: 'oct.-22', cr: 181000, va: 181500, vp: 181500 },
  { month: 'nov.-22', cr: 261000, va: 248800, vp: 252000 },
  { month: 'déc.-22', cr: 335000, va: 333000, vp: 321000 },
  { month: 'janv.-23', cr: 355000, va: 340400, vp: 344000 },
  { month: 'févr.-23', cr: 357000, va: 341400, vp: 346000 }
];

export const SuiviBudgetMatrix: React.FC<SuiviBudgetMatrixProps> = ({
  projectId,
  fileId,
  currentDossierId,
  activeSubTab = 'rapport'
}) => {
  const dossierSubKey = currentDossierId ? `-${currentDossierId}` : '';
  const storageKey = `faab-budget-matrix-${projectId}-${fileId}${dossierSubKey}`;

  const [monthlyData, setMonthlyData] = useState<MonthlyData[]>(() => {
    const saved = localStorage.getItem(storageKey);
    if (saved) {
      try { return JSON.parse(saved); } catch (e) { }
    }
    return DEFAULT_MONTHLY;
  });

  const [baa, setBaa] = useState<number>(346000);
  const [currentTab, setCurrentTab] = useState<string>(activeSubTab || 'rapport');

  useEffect(() => {
    if (activeSubTab) setCurrentTab(activeSubTab);
  }, [activeSubTab]);

  useEffect(() => {
    localStorage.setItem(storageKey, JSON.stringify(monthlyData));
  }, [monthlyData, storageKey]);

  // Recalcul automatique des métriques Earned Value pour chaque mois
  const computedMetrics = useMemo(() => {
    return monthlyData.map(d => {
      const ec = d.va - d.cr; // Écart Coût
      const ed = d.va - d.vp; // Écart Délais
      const ipc = d.cr > 0 ? d.va / d.cr : 1; // CPI
      const ipd = d.vp > 0 ? d.va / d.vp : 1; // SPI
      const cfe = ipc > 0 ? Math.round(baa / ipc) : baa; // Coût Final Estimé
      return {
        ...d,
        ec,
        ed,
        ipc: parseFloat(ipc.toFixed(2)),
        ipd: parseFloat(ipd.toFixed(2)),
        cfe
      };
    });
  }, [monthlyData, baa]);

  const lastMetric = computedMetrics[computedMetrics.length - 1];

  const handleCellChange = (index: number, field: 'cr' | 'va' | 'vp', value: string) => {
    const num = parseFloat(value) || 0;
    const next = [...monthlyData];
    next[index] = { ...next[index], [field]: num };
    setMonthlyData(next);
  };

  return (
    <div style={{ flex: 1, display: 'flex', flexDirection: 'column', background: 'var(--main-bg)', color: 'var(--main-text)', overflowY: 'auto' }}>
      {/* Header Fiche Projet */}
      <div style={{ padding: '1.2rem 2rem', background: 'var(--main-surface)', borderBottom: '1px solid var(--main-border)', display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
        <div>
          <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
            <span style={{ fontSize: '1.2rem' }}>💰</span>
            <h2 style={{ margin: 0, fontSize: '1.25rem', fontWeight: 800 }}>28 — Suivi Budgétaire & Earned Value (VA)</h2>
          </div>
          <p style={{ margin: '4px 0 0', color: 'var(--text-muted)', fontSize: '0.85rem' }}>
            Calculs automatiques Valeur Planifiée (VP), Coût Réel (CR), Valeur Acquise (VA), et Indices de Performance
          </p>
        </div>

        {/* KPIs rapides */}
        <div style={{ display: 'flex', gap: '1rem' }}>
          <div style={{ padding: '6px 14px', borderRadius: '10px', background: 'rgba(16, 185, 129, 0.1)', border: '1px solid rgba(16, 185, 129, 0.25)', textAlign: 'center' }}>
            <div style={{ fontSize: '0.7rem', color: '#10b981', fontWeight: 700, textTransform: 'uppercase' }}>IPC (Coûts)</div>
            <div style={{ fontSize: '1.2rem', fontWeight: 800, color: lastMetric.ipc >= 1 ? '#10b981' : '#ef4444' }}>{lastMetric.ipc}</div>
          </div>
          <div style={{ padding: '6px 14px', borderRadius: '10px', background: 'rgba(56, 189, 248, 0.1)', border: '1px solid rgba(56, 189, 248, 0.25)', textAlign: 'center' }}>
            <div style={{ fontSize: '0.7rem', color: '#38bdf8', fontWeight: 700, textTransform: 'uppercase' }}>IPD (Délais)</div>
            <div style={{ fontSize: '1.2rem', fontWeight: 800, color: lastMetric.ipd >= 1 ? '#10b981' : '#f59e0b' }}>{lastMetric.ipd}</div>
          </div>
          <div style={{ padding: '6px 14px', borderRadius: '10px', background: 'rgba(168, 85, 247, 0.1)', border: '1px solid rgba(168, 85, 247, 0.25)', textAlign: 'center' }}>
            <div style={{ fontSize: '0.7rem', color: '#a855f7', fontWeight: 700, textTransform: 'uppercase' }}>Coût Final Estimé</div>
            <div style={{ fontSize: '1.2rem', fontWeight: 800, color: '#f8fafc' }}>{lastMetric.cfe.toLocaleString()} €</div>
          </div>
        </div>
      </div>

      {/* Barre de navigation d'onglets interne */}
      <div style={{ display: 'flex', gap: '8px', padding: '0.8rem 2rem', background: 'rgba(15, 23, 42, 0.3)', borderBottom: '1px solid var(--main-border)' }}>
        {[
          { id: 'rapport', label: '📊 Matrice VA & Courbe en S' },
          { id: 'vp', label: '📈 Valeur Planifiée (VP)' },
          { id: 'cr', label: '💳 Coût Réel (CR)' },
          { id: 'va', label: '🟢 Valeur Acquise (VA)' },
          { id: 'avanc', label: '📈 Graphique Avancement' }
        ].map(tab => (
          <button
            key={tab.id}
            onClick={() => setCurrentTab(tab.id)}
            style={{
              background: currentTab === tab.id ? '#10b981' : 'rgba(255, 255, 255, 0.05)',
              color: currentTab === tab.id ? '#fff' : '#94a3b8',
              border: '1px solid rgba(255, 255, 255, 0.1)',
              padding: '0.45rem 0.9rem',
              borderRadius: '8px',
              fontSize: '0.82rem',
              fontWeight: 600,
              cursor: 'pointer',
              transition: 'all 0.15s ease'
            }}
          >
            {tab.label}
          </button>
        ))}
      </div>

      {/* Contenu de la vue active */}
      <div style={{ padding: '1.5rem 2rem', flex: 1 }}>
        {/* VUE 1 : RAPPORT VA CONSOLIDÉ (LA VRAIE MATRICE EXCEL) */}
        {(currentTab === 'rapport' || currentTab === 'option4' || currentTab === 'option1') && (
          <div>
            {/* Graphique Courbe en S Interactif Recharts */}
            <div style={{
              background: 'var(--main-surface)',
              borderRadius: '14px',
              border: '1px solid var(--main-border)',
              padding: '1.2rem 1.5rem',
              marginBottom: '1.5rem',
              boxShadow: '0 4px 16px rgba(0,0,0,0.1)'
            }}>
              <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '1rem' }}>
                <h3 style={{ margin: 0, fontSize: '1rem', fontWeight: 700, color: '#f8fafc' }}>
                  📈 S-Curve : Courbe en S de l'Avancement Budgétaire (VP vs VA vs CR)
                </h3>
                <span style={{ fontSize: '0.78rem', color: '#10b981', fontWeight: 600 }}>
                  ● Mise à jour automatique en temps réel à chaque modification
                </span>
              </div>
              <div style={{ height: '280px', width: '100%' }}>
                <ResponsiveContainer width="100%" height="100%">
                  <LineChart data={computedMetrics} margin={{ top: 10, right: 30, left: 20, bottom: 5 }}>
                    <CartesianGrid strokeDasharray="3 3" stroke="rgba(255,255,255,0.06)" />
                    <XAxis dataKey="month" stroke="#94a3b8" fontSize={11} />
                    <YAxis stroke="#94a3b8" fontSize={11} tickFormatter={(val) => `${Math.round(val / 1000)}k€`} />
                    <Tooltip
                      formatter={(val: any) => [`${Number(val).toLocaleString()} €`]}
                      contentStyle={{ background: '#0f172a', border: '1px solid #334155', borderRadius: '8px', color: '#f8fafc' }}
                    />
                    <Legend />
                    <Line type="monotone" dataKey="vp" name="Valeur Planifiée (VP)" stroke="#38bdf8" strokeWidth={3} dot={{ r: 4 }} activeDot={{ r: 6 }} />
                    <Line type="monotone" dataKey="va" name="Valeur Acquise (VA)" stroke="#10b981" strokeWidth={3} dot={{ r: 4 }} activeDot={{ r: 6 }} />
                    <Line type="monotone" dataKey="cr" name="Coût Réel (CR)" stroke="#f59e0b" strokeWidth={3} strokeDasharray="5 5" dot={{ r: 4 }} />
                  </LineChart>
                </ResponsiveContainer>
              </div>
            </div>

            {/* Tableau Matrice Consolidée */}
            <div style={{ background: 'var(--main-surface)', borderRadius: '14px', border: '1px solid var(--main-border)', overflowX: 'auto', boxShadow: '0 4px 16px rgba(0,0,0,0.1)' }}>
              <div style={{ padding: '1rem 1.5rem', background: 'rgba(16, 185, 129, 0.1)', borderBottom: '1px solid var(--main-border)', display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                <span style={{ fontWeight: 800, fontSize: '0.9rem', color: '#10b981' }}>
                  TABLEAU DE PERFORMANCE DE LA VALEUR ACQUISE (12 MOIS)
                </span>
                <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
                  <span style={{ fontSize: '0.8rem', color: 'var(--text-muted)' }}>Budget à l'achèvement (BAA) :</span>
                  <input
                    type="number"
                    value={baa}
                    onChange={(e) => setBaa(parseFloat(e.target.value) || 0)}
                    style={{ background: 'rgba(0,0,0,0.3)', border: '1px solid rgba(255,255,255,0.2)', color: '#fff', borderRadius: '6px', padding: '4px 8px', fontSize: '0.85rem', width: '110px', textAlign: 'right' }}
                  />
                  <span style={{ fontSize: '0.8rem', color: '#fff' }}>€</span>
                </div>
              </div>

              <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '0.82rem' }}>
                <thead>
                  <tr style={{ background: 'rgba(15, 23, 42, 0.8)', borderBottom: '1px solid var(--main-border)' }}>
                    <th style={{ padding: '12px 14px', textAlign: 'left', minWidth: '220px', color: '#f8fafc', position: 'sticky', left: 0, background: '#0f172a' }}>Métrique / Période</th>
                    {MONTHS.map(m => (
                      <th key={m} style={{ padding: '12px 10px', textAlign: 'right', minWidth: '85px', color: '#94a3b8' }}>{m}</th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {/* Ligne 1 : CR cumulé */}
                  <tr style={{ borderBottom: '1px solid var(--main-border)' }}>
                    <td style={{ padding: '10px 14px', fontWeight: 700, color: '#f59e0b', position: 'sticky', left: 0, background: 'var(--main-surface)' }}>
                      Coût Réel cumulé (CR)
                    </td>
                    {monthlyData.map((d, i) => (
                      <td key={d.month} style={{ padding: '6px 8px', textAlign: 'right' }}>
                        <input
                          type="number"
                          value={d.cr}
                          onChange={(e) => handleCellChange(i, 'cr', e.target.value)}
                          style={{ width: '100%', background: 'transparent', border: 'none', color: '#f8fafc', textAlign: 'right', fontSize: '0.82rem', fontWeight: 600 }}
                        />
                      </td>
                    ))}
                  </tr>

                  {/* Ligne 2 : VA cumulée */}
                  <tr style={{ borderBottom: '1px solid var(--main-border)', background: 'rgba(16, 185, 129, 0.03)' }}>
                    <td style={{ padding: '10px 14px', fontWeight: 700, color: '#10b981', position: 'sticky', left: 0, background: 'var(--main-surface)' }}>
                      Valeur acquise cumulée (VA)
                    </td>
                    {monthlyData.map((d, i) => (
                      <td key={d.month} style={{ padding: '6px 8px', textAlign: 'right' }}>
                        <input
                          type="number"
                          value={d.va}
                          onChange={(e) => handleCellChange(i, 'va', e.target.value)}
                          style={{ width: '100%', background: 'transparent', border: 'none', color: '#10b981', textAlign: 'right', fontSize: '0.82rem', fontWeight: 700 }}
                        />
                      </td>
                    ))}
                  </tr>

                  {/* Ligne 3 : VP cumulée */}
                  <tr style={{ borderBottom: '1px solid var(--main-border)' }}>
                    <td style={{ padding: '10px 14px', fontWeight: 700, color: '#38bdf8', position: 'sticky', left: 0, background: 'var(--main-surface)' }}>
                      Valeur Planifiée cumulée (VP)
                    </td>
                    {monthlyData.map((d, i) => (
                      <td key={d.month} style={{ padding: '6px 8px', textAlign: 'right' }}>
                        <input
                          type="number"
                          value={d.vp}
                          onChange={(e) => handleCellChange(i, 'vp', e.target.value)}
                          style={{ width: '100%', background: 'transparent', border: 'none', color: '#38bdf8', textAlign: 'right', fontSize: '0.82rem', fontWeight: 600 }}
                        />
                      </td>
                    ))}
                  </tr>

                  {/* Ligne 4 : Écart de Coût (EC = VA - CR) */}
                  <tr style={{ borderBottom: '1px solid var(--main-border)', background: 'rgba(0,0,0,0.1)' }}>
                    <td style={{ padding: '10px 14px', fontWeight: 800, position: 'sticky', left: 0, background: 'var(--main-surface)' }}>
                      Écart de Coût (EC = VA - CR)
                    </td>
                    {computedMetrics.map(m => (
                      <td key={m.month} style={{ padding: '10px 10px', textAlign: 'right', fontWeight: 700, color: m.ec >= 0 ? '#10b981' : '#ef4444' }}>
                        {m.ec.toLocaleString()} €
                      </td>
                    ))}
                  </tr>

                  {/* Ligne 5 : Écart de Délais (ED = VA - VP) */}
                  <tr style={{ borderBottom: '1px solid var(--main-border)', background: 'rgba(0,0,0,0.1)' }}>
                    <td style={{ padding: '10px 14px', fontWeight: 800, position: 'sticky', left: 0, background: 'var(--main-surface)' }}>
                      Écart de Délais (ED = VA - VP)
                    </td>
                    {computedMetrics.map(m => (
                      <td key={m.month} style={{ padding: '10px 10px', textAlign: 'right', fontWeight: 700, color: m.ed >= 0 ? '#10b981' : '#f59e0b' }}>
                        {m.ed.toLocaleString()} €
                      </td>
                    ))}
                  </tr>

                  {/* Ligne 6 : IPC (VA / CR) */}
                  <tr style={{ borderBottom: '1px solid var(--main-border)' }}>
                    <td style={{ padding: '10px 14px', fontWeight: 800, position: 'sticky', left: 0, background: 'var(--main-surface)' }}>
                      Indice Performance Coûts (IPC)
                    </td>
                    {computedMetrics.map(m => (
                      <td key={m.month} style={{ padding: '10px 10px', textAlign: 'right', fontWeight: 800, color: m.ipc >= 1 ? '#10b981' : '#ef4444' }}>
                        {m.ipc.toFixed(2)}
                      </td>
                    ))}
                  </tr>

                  {/* Ligne 7 : IPD (VA / VP) */}
                  <tr style={{ borderBottom: '1px solid var(--main-border)' }}>
                    <td style={{ padding: '10px 14px', fontWeight: 800, position: 'sticky', left: 0, background: 'var(--main-surface)' }}>
                      Indice Performance Délais (IPD)
                    </td>
                    {computedMetrics.map(m => (
                      <td key={m.month} style={{ padding: '10px 10px', textAlign: 'right', fontWeight: 800, color: m.ipd >= 1 ? '#10b981' : '#f59e0b' }}>
                        {m.ipd.toFixed(2)}
                      </td>
                    ))}
                  </tr>

                  {/* Ligne 8 : Coût Final Estimé (CFE = BAA / IPC) */}
                  <tr style={{ background: 'rgba(16, 185, 129, 0.08)' }}>
                    <td style={{ padding: '12px 14px', fontWeight: 800, color: '#f8fafc', position: 'sticky', left: 0, background: '#0f172a' }}>
                      Coût Final Estimé (CFE)
                    </td>
                    {computedMetrics.map(m => (
                      <td key={m.month} style={{ padding: '12px 10px', textAlign: 'right', fontWeight: 800, color: '#f8fafc' }}>
                        {m.cfe.toLocaleString()} €
                      </td>
                    ))}
                  </tr>
                </tbody>
              </table>
            </div>

            <div style={{ marginTop: '1rem', display: 'flex', gap: '2rem', fontSize: '0.78rem', color: 'var(--text-muted)' }}>
              <span>🟢 Si IPC &gt; 1 : Budget sous-consommé (Économies)</span>
              <span>🔴 Si IPC &lt; 1 : Budget sur-consommé (Dépassement)</span>
              <span>🟢 Si IPD &gt; 1 : En avance sur le planning</span>
              <span>🟡 Si IPD &lt; 1 : Retard sur le planning</span>
            </div>
          </div>
        )}

        {/* VUE 2 : VALEUR PLANIFIÉE (VP) DÉTAILLÉE PAR TÂCHES */}
        {currentTab === 'vp' && (
          <div style={{ background: 'var(--main-surface)', borderRadius: '14px', border: '1px solid var(--main-border)', padding: '1.5rem' }}>
            <h3 style={{ margin: '0 0 1rem', fontSize: '1.1rem', fontWeight: 800, color: '#38bdf8' }}>
              📈 Valeur Planifiée (VP) par Lot de Travail
            </h3>
            <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '0.85rem' }}>
              <thead>
                <tr style={{ background: 'rgba(56, 189, 248, 0.1)', borderBottom: '1px solid var(--main-border)' }}>
                  <th style={{ padding: '10px 14px', textAlign: 'left' }}># WBS</th>
                  <th style={{ padding: '10px 14px', textAlign: 'left' }}>Lot de Travail / Tâche</th>
                  <th style={{ padding: '10px 14px', textAlign: 'right' }}>Budget Alloué</th>
                  <th style={{ padding: '10px 14px', textAlign: 'right' }}>Valeur Planifiée à Date</th>
                  <th style={{ padding: '10px 14px', textAlign: 'center' }}>Statut</th>
                </tr>
              </thead>
              <tbody>
                {[
                  { wbs: '1.1', task: "Étude d'opportunité & Cadrage", budget: 19500, vp: 19500, status: 'Planifié' },
                  { wbs: '1.2', task: 'Expression des besoins & Exigences', budget: 36000, vp: 36000, status: 'Planifié' },
                  { wbs: '2.1', task: 'Architecture & Spécifications', budget: 55000, vp: 55000, status: 'Planifié' },
                  { wbs: '2.2', task: 'Développement Lot 1 Finance', budget: 120000, vp: 110000, status: 'Planifié' },
                  { wbs: '3.1', task: 'Tests d’intégration & Recette', budget: 45000, vp: 35000, status: 'Planifié' },
                  { wbs: '4.1', task: 'Déploiement & Conduite du changement', budget: 70500, vp: 25000, status: 'Planifié' }
                ].map((row, idx) => (
                  <tr key={idx} style={{ borderBottom: '1px solid var(--main-border)' }}>
                    <td style={{ padding: '10px 14px', fontWeight: 700, color: '#38bdf8' }}>{row.wbs}</td>
                    <td style={{ padding: '10px 14px', fontWeight: 600 }}>{row.task}</td>
                    <td style={{ padding: '10px 14px', textAlign: 'right', fontWeight: 700 }}>{row.budget.toLocaleString()} €</td>
                    <td style={{ padding: '10px 14px', textAlign: 'right', fontWeight: 700, color: '#38bdf8' }}>{row.vp.toLocaleString()} €</td>
                    <td style={{ padding: '10px 14px', textAlign: 'center' }}>
                      <span style={{ padding: '3px 10px', borderRadius: '20px', background: 'rgba(56, 189, 248, 0.15)', color: '#38bdf8', fontSize: '0.75rem', fontWeight: 700 }}>
                        {row.status}
                      </span>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}

        {/* VUE 3 : COÛT RÉEL (CR) DÉTAILLÉ */}
        {currentTab === 'cr' && (
          <div style={{ background: 'var(--main-surface)', borderRadius: '14px', border: '1px solid var(--main-border)', padding: '1.5rem' }}>
            <h3 style={{ margin: '0 0 1rem', fontSize: '1.1rem', fontWeight: 800, color: '#f59e0b' }}>
              💳 Coûts Réels Constatés (CR) par Catégorie de Dépenses
            </h3>
            <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '0.85rem' }}>
              <thead>
                <tr style={{ background: 'rgba(245, 158, 11, 0.1)', borderBottom: '1px solid var(--main-border)' }}>
                  <th style={{ padding: '10px 14px', textAlign: 'left' }}>Poste de Dépense</th>
                  <th style={{ padding: '10px 14px', textAlign: 'left' }}>Responsable / Fournisseur</th>
                  <th style={{ padding: '10px 14px', textAlign: 'right' }}>Coût Engagé</th>
                  <th style={{ padding: '10px 14px', textAlign: 'right' }}>Facturé & Payé</th>
                  <th style={{ padding: '10px 14px', textAlign: 'center' }}>Écart vs Budget</th>
                </tr>
              </thead>
              <tbody>
                {[
                  { poste: 'Prestations Régie Externes (Dev)', resp: 'Consulting IT', engage: 145000, paye: 142000, ecart: '+2 000 €' },
                  { poste: 'Licences & Infrastructure Cloud', resp: 'Fournisseur Cloud', engage: 65000, paye: 64500, ecart: '-500 €' },
                  { poste: 'Formations & Conduite changement', resp: 'RH Formation', engage: 32000, paye: 28000, ecart: '0 €' },
                  { poste: 'Matériel & Équipements tests', resp: 'Achats Tech', engage: 18500, paye: 18500, ecart: '+500 €' }
                ].map((row, idx) => (
                  <tr key={idx} style={{ borderBottom: '1px solid var(--main-border)' }}>
                    <td style={{ padding: '10px 14px', fontWeight: 700 }}>{row.poste}</td>
                    <td style={{ padding: '10px 14px', color: 'var(--text-muted)' }}>{row.resp}</td>
                    <td style={{ padding: '10px 14px', textAlign: 'right', fontWeight: 700 }}>{row.engage.toLocaleString()} €</td>
                    <td style={{ padding: '10px 14px', textAlign: 'right', fontWeight: 700, color: '#f59e0b' }}>{row.paye.toLocaleString()} €</td>
                    <td style={{ padding: '10px 14px', textAlign: 'center', fontWeight: 700, color: row.ecart.startsWith('+') ? '#ef4444' : '#10b981' }}>
                      {row.ecart}
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}

        {/* VUE 4 : VALEUR ACQUISE (VA) */}
        {currentTab === 'va' && (
          <div style={{ background: 'var(--main-surface)', borderRadius: '14px', border: '1px solid var(--main-border)', padding: '1.5rem' }}>
            <h3 style={{ margin: '0 0 1rem', fontSize: '1.1rem', fontWeight: 800, color: '#10b981' }}>
              🟢 Valeur Acquise (VA) - Avancement Physique Réel
            </h3>
            <p style={{ color: 'var(--text-muted)', fontSize: '0.85rem', marginBottom: '1.2rem' }}>
              La Valeur Acquise représente le budget valorisé correspondant aux livrables réellement terminés et validés par le comité de pilotage.
            </p>
            <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(280px, 1fr))', gap: '1rem' }}>
              {[
                { phase: 'Phase 1 : Cadrage', pct: 100, val: 55500, total: 55500 },
                { phase: 'Phase 2 : Conception', pct: 100, val: 120000, total: 120000 },
                { phase: 'Phase 3 : Réalisation', pct: 85, val: 125000, total: 147000 },
                { phase: 'Phase 4 : Clôture & REX', pct: 40, val: 9400, total: 23500 }
              ].map((item, i) => (
                <div key={i} style={{ background: 'rgba(0,0,0,0.2)', padding: '1.2rem', borderRadius: '10px', border: '1px solid var(--main-border)' }}>
                  <div style={{ display: 'flex', justifyContent: 'space-between', marginBottom: '6px' }}>
                    <span style={{ fontWeight: 700, fontSize: '0.9rem' }}>{item.phase}</span>
                    <span style={{ fontWeight: 800, color: '#10b981' }}>{item.pct}%</span>
                  </div>
                  <div style={{ height: '8px', background: 'rgba(255,255,255,0.1)', borderRadius: '4px', overflow: 'hidden', marginBottom: '8px' }}>
                    <div style={{ width: `${item.pct}%`, height: '100%', background: '#10b981' }}></div>
                  </div>
                  <div style={{ display: 'flex', justifyContent: 'space-between', fontSize: '0.75rem', color: 'var(--text-muted)' }}>
                    <span>Acquis : {item.val.toLocaleString()} €</span>
                    <span>Total : {item.total.toLocaleString()} €</span>
                  </div>
                </div>
              ))}
            </div>
          </div>
        )}

        {/* VUE 5 : AVANCEMENT CONSOLIDÉ */}
        {currentTab === 'avanc' && (
          <div style={{ background: 'var(--main-surface)', borderRadius: '14px', border: '1px solid var(--main-border)', padding: '1.5rem' }}>
            <h3 style={{ margin: '0 0 1rem', fontSize: '1.1rem', fontWeight: 800, color: '#f8fafc' }}>
              📈 Synthèse d'Avancement Prévisionnel
            </h3>
            <p style={{ color: 'var(--text-muted)', fontSize: '0.85rem' }}>
              Modélisation des tendances et projection de fin de projet. Les calculs indiquent un atterrissage budgétaire maîtrisé avec un IPC moyen de {lastMetric.ipc}.
            </p>
          </div>
        )}
      </div>
    </div>
  );
};
