import { tabsData } from './data';
import type { ProjectFile } from './data';
import React, { useState, useEffect, useMemo, Fragment } from 'react';
import { Cockpit } from './Cockpit';
import { SuiviBudgetMatrix } from './modules/SuiviBudgetMatrix';
import { TableauDeBordMatrix } from './modules/TableauDeBordMatrix';
import { RapportFlashMatrix } from './modules/RapportFlashMatrix';
import { BilanProjetMatrix } from './modules/BilanProjetMatrix';
import './App.css';

// ─── Global data version: bump this string to reset ALL cached table data ───
const APP_DATA_VERSION = '2026-03-06-v2';
(() => {
  const VK = '__faab_data_version__';
  if (localStorage.getItem(VK) !== APP_DATA_VERSION) {
    // Keep portfolios, wipe everything else
    const portfolioKey = Object.keys(localStorage).find(k => k.startsWith('portfolios'));
    const portfolioValue = portfolioKey ? localStorage.getItem(portfolioKey) : null;
    localStorage.clear();
    if (portfolioKey && portfolioValue) localStorage.setItem(portfolioKey, portfolioValue);
    localStorage.setItem(VK, APP_DATA_VERSION);
  }
})();

import { ProjectReportGenerator } from './modules/ProjectReportGenerator';
import { ProjectTrainingViewer } from './modules/ProjectTrainingViewer';

const ExcelIcon = () => (
  <svg xmlns="http://www.w3.org/2000/svg" width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="#10b981" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
    <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"></path>
    <polyline points="14 2 14 8 20 8"></polyline>
    <path d="M8 13h2"></path>
    <path d="M8 17h2"></path>
    <path d="M14 13h2"></path>
    <path d="M14 17h2"></path>
  </svg>
);

const WordIcon = () => (
  <svg xmlns="http://www.w3.org/2000/svg" width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="#38bdf8" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
    <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"></path>
    <polyline points="14 2 14 8 20 8"></polyline>
    <line x1="16" y1="13" x2="8" y2="13"></line>
    <line x1="16" y1="17" x2="8" y2="17"></line>
    <polyline points="10 9 9 9 8 9"></polyline>
  </svg>
);

const PptIcon = () => (
  <svg xmlns="http://www.w3.org/2000/svg" width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="#f59e0b" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
    <rect x="2" y="3" width="20" height="14" rx="2" ry="2"></rect>
    <line x1="8" y1="21" x2="16" y2="21"></line>
    <line x1="12" y1="17" x2="12" y2="21"></line>
    <path d="M7 8h4a2 2 0 0 1 0 4H7z"></path>
  </svg>
);

// ─── Utils ─────────────────────────────────────────────────────────
const exportToCSV = (data: any[], filename: string) => {
  if (!data || data.length === 0) return;
  const headers = Object.keys(data[0]);
  const csvData = [
    headers.join(';'),
    ...data.map(row => headers.map(header => {
      let cell = row[header] === null || row[header] === undefined ? '' : row[header];
      cell = cell.toString().replace(/"/g, '""');
      if (cell.includes(';') || cell.includes('\n') || cell.includes('"')) {
        cell = `"${cell}"`;
      }
      return cell;
    }).join(';'))
  ].join('\r\n');

  const blob = new Blob([new Uint8Array([0xEF, 0xBB, 0xBF]), csvData], { type: 'text/csv;charset=utf-8;' });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.setAttribute("href", url);
  link.setAttribute("download", filename);
  link.style.visibility = 'hidden';
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
};

interface Project {
  id: string;
  name: string;
}

interface Portfolio {
  id: string;
  name: string;
  projects: Project[];
}

interface SelectionDossier {
  id: string;
  name: string;
  createdAt: number;
}

const getFileOptions = (fileId?: string) => {
  let options: { id: string; label: string }[] = [
    { id: 'option1', label: '➡️ Option 1' },
    { id: 'option2', label: '➡️ Option 2' },
  ];

  if (fileId === 'selection-projets') {
    options = [
      { id: 'option_liste', label: '📋 Liste' },
      { id: 'option1', label: '📊 Matrice' },
      { id: 'option2', label: '📑 Référentiel' },
      { id: 'option_dashboard', label: '📉 Tableau de bord' },
      { id: 'option_dossiers', label: '📁 Dossiers de Sélection' },
      { id: 'option_settings', label: '⚙️ Paramétrage' }
    ];
  } else if (fileId === 'matrice-faisabilite') {
    options = [
      { id: 'option_liste', label: '📋 Liste' },
      { id: 'option1', label: '📊 Matrice' },
      { id: 'option2', label: '📑 Référentiel' },
      { id: 'option_dashboard', label: '📉 Tableau de bord' },
      { id: 'option_dossiers', label: '📁 Dossiers de Sélection' },
      { id: 'option_settings', label: '⚙️ Paramétrage' }
    ];
  } else if (fileId === 'prenantes-1') {
    options = [
      { id: 'option1', label: '📋 Registre' },
      { id: 'option2', label: '🌍 Matrice' },
      { id: 'option_dossiers', label: '📁 Dossiers' }
    ];
  } else if (fileId === 'prenantes-2') {
    options = [
      { id: 'option1', label: '📋 Cartographie' },
      { id: 'option2', label: '📈 Graphique' },
      { id: 'option_dossiers', label: '📁 Dossiers' }
    ];
  } else if (fileId === 'macro-planning') {
    options = [
      { id: 'option1', label: '📊 Table' },
      { id: 'option2', label: '📅 Graphique' },
      { id: 'option_pert', label: '🕸️ G Perte' },
      { id: 'option3', label: '👥 Ressources' },
      { id: 'option_dossiers', label: '📁 Dossiers' },
      { id: 'option4', label: '⚙️ Paramétrage' }
    ];
  } else if (fileId === 'risques-1') {
    options = [
      { id: 'option1', label: 'ℹ️ Guide' },
      { id: 'option2', label: '📏 Echelles' },
      { id: 'option3', label: '📋 Registre Risques' },
      { id: 'option4', label: '⚠️ Registre Problèmes' },
      { id: 'option5', label: '📊 Matrice' },
      { id: 'option_dossiers', label: '📁 Dossiers' },
      { id: 'option6', label: '⚙️ Paramétrage' }
    ];
  } else if (fileId === 'risques-2') {
    options = [
      { id: 'option1', label: '📋 Cartographie' },
      { id: 'option2', label: '📈 Dashboard' },
      { id: 'option_dossiers', label: '📁 Dossiers' },
      { id: 'option3', label: '⚙️ Paramétrage' }
    ];
  } else if (fileId === 'comm-fr' || fileId === 'comm-act-en') {
    options = [
      { id: 'option1', label: '📢 Plan d\'actions' },
      { id: 'option2', label: '🎯 Sponsor' },
      { id: 'option_dossiers', label: '📁 Dossiers' }
    ];
  } else if (fileId === 'competences') {
    options = [
      { id: 'option1', label: '🔢 Matrice' },
      { id: 'option2', label: '📊 Comparaison' },
      { id: 'option3', label: '📉 Dashboard' },
      { id: 'option_dossiers', label: '📁 Dossiers' },
      { id: 'option4', label: '⚙️ Paramètre' }
    ];
  } else if (fileId === 'wbs') {
    options = [
      { id: 'option1', label: '🌳 PBS (Livrables)' },
      { id: 'option2', label: '🌳 WBS (Tâches)' },
      { id: 'option_dossiers', label: '📁 Dossiers' }
    ];
  } else if (fileId === 'raci') {
    options = [
      { id: 'option1', label: '📊 RACI Complet' },
      { id: 'option2', label: '✅ R : Réalisateur' },
      { id: 'option3', label: '📋 A : Approbateur' },
      { id: 'option4', label: '💡 C/I' },
      { id: 'option_dossiers', label: '📁 Dossiers' }
    ];
  } else if (fileId === 'gantt') {
    options = [
      { id: 'option1', label: '📊 Table' },
      { id: 'option2', label: '📅 Graphique' },
      { id: 'option3', label: '👥 Ressources' },
      { id: 'option_dossiers', label: '📁 Dossiers' },
      { id: 'option4', label: '⚙️ Paramétrage' }
    ];
  } else if (fileId === 'plan-charge') {
    options = [
      { id: 'option1', label: '📖 Guide' },
      { id: 'option2', label: '📊 Planning' },
      { id: 'option3', label: '👥 Ressources' },
      { id: 'option_dossiers', label: '📁 Dossiers' },
      { id: 'option4', label: '⚙️ Paramétrage' }
    ];
  } else if (fileId === 'matrice-exigences') {
    options = [
      { id: 'option1', label: '📋 Matrice' },
      { id: 'option_dossiers', label: '📁 Dossiers' }
    ];
  } else if (fileId === 'suivi-actions') {
    options = [
      { id: 'option1', label: '✅ Liste' },
      { id: 'option_dossiers', label: '📁 Dossiers' },
      { id: 'option2', label: '⚙️ Paramétrage' }
    ];
  } else if (fileId === 'suivi-livrables') {
    options = [
      { id: 'option1', label: '📦 Liste' },
      { id: 'option_dossiers', label: '📁 Dossiers' },
      { id: 'option2', label: '⚙️ Paramétrage' }
    ];
  } else if (fileId === 'chg-req-reg') {
    options = [
      { id: 'option1', label: '📝 Registre' },
      { id: 'option_dossiers', label: '📁 Dossiers' },
      { id: 'option2', label: '⚙️ Configuration' }
    ];
  } else if (fileId === 'amdec') {
    options = [
      { id: 'option1', label: '📏 Échelles' },
      { id: 'option2', label: '💻 AMDEC Logiciel' },
      { id: 'option3', label: '🏗️ AMDEC Matériel' },
      { id: 'option4', label: '📊 Répartition' },
      { id: 'option_dossiers', label: '📁 Dossiers' }
    ];
  } else if (fileId === 'plan-test') {
    options = [
      { id: 'option1', label: '🧪 Tests' },
      { id: 'option2', label: '📋 Consolidation' },
      { id: 'option_dossiers', label: '📁 Dossiers' }
    ];
  } else if (fileId === 'suivi-budget') {
    options = [
      { id: 'option1', label: '📊 Rapport VA & S-Curve' },
      { id: 'option2', label: '💰 Coût Réel (CR)' },
      { id: 'option3', label: '🟢 Valeur Acquise (VA)' },
      { id: 'option4', label: '📈 Valeur Planifiée (VP)' },
      { id: 'option_dossiers', label: '📁 Dossiers' }
    ];
  } else if (fileId === 'tb') {
    options = [
      { id: 'option1', label: '📊 Dashboard Consolidé' },
      { id: 'option2', label: '📈 Écarts Délais & Coûts' },
      { id: 'option3', label: '💶 Coûts Consommés' },
      { id: 'option4', label: '📋 Tâches & Phases' },
      { id: 'option_dossiers', label: '📁 Dossiers' }
    ];
  } else if (fileId === 'rapport-flash-xls') {
    options = [
      { id: 'option1', label: '⚡ Rapport Flash Hebdo' },
      { id: 'option_dossiers', label: '📁 Dossiers' }
    ];
  } else if (fileId === 'bilan') {
    options = [
      { id: 'option1', label: '📋 Bilan & Écarts' },
      { id: 'option2', label: '💡 REX 9 Domaines' },
      { id: 'option3', label: '⭐ Matrice Satisfaction' },
      { id: 'option4', label: '📦 Dépôt & Clôture' },
      { id: 'option_dossiers', label: '📁 Dossiers' }
    ];
  }
  return options;
};

const HeaderDropdownOptions = ({ onSelectOption, fileId, activePage }: { onSelectOption: (option: string) => void; fileId?: string; activePage?: string | null }) => {
  const options = getFileOptions(fileId);

  return (
    <div style={{ display: 'flex', gap: '0.6rem', flexWrap: 'wrap', alignItems: 'center' }}>
      {options.map((opt) => (
        <button
          key={opt.id}
          className={`btn-option-pill ${activePage === opt.id || (!activePage && opt.id === 'option1') ? 'active' : ''}`}
          onClick={() => onSelectOption(opt.id)}
        >
          {opt.label}
        </button>
      ))}
    </div>
  );
};

const TemplateControls = ({ fileId, projectId, onExistantClick }: { fileId: string, projectId: string, onExistantClick?: () => void }) => {
  const isSelection = fileId === 'selection-projets';

  const handleNew = () => {
    const keysToClear = [];
    for (let i = 0; i < localStorage.length; i++) {
      const key = localStorage.key(i);
      if (key && (key.includes(fileId) || key.includes(`tab-data-${projectId}-${fileId}`))) {
        keysToClear.push(key);
      }
    }
    keysToClear.forEach(k => localStorage.removeItem(k));
    window.location.reload();
  };

  const handleExistant = () => {
    if (onExistantClick) {
      onExistantClick();
    } else {
      alert("Fonctionnalité 'Existant' : Choix d'un template déjà enregistré. (Bientôt disponible)");
    }
  };

  return (
    <div style={{ display: 'flex', gap: '0.5rem', marginLeft: '1rem', borderLeft: '1px solid var(--border)', paddingLeft: '1rem' }}>
      {true && (
        <button
          onClick={handleNew}
          style={{
            background: 'rgba(16, 185, 129, 0.2)',
            color: '#34d399',
            border: '1px solid rgba(16, 185, 129, 0.3)',
            padding: '6px 12px',
            borderRadius: '8px',
            fontSize: '0.8rem',
            fontWeight: 700,
            cursor: 'pointer'
          }}
        >
          ✨ Nouveau
        </button>
      )}
      <button
        onClick={handleExistant}
        style={{
          background: 'rgba(59, 130, 246, 0.2)',
          color: '#60a5fa',
          border: '1px solid rgba(59, 130, 246, 0.3)',
          padding: '6px 12px',
          borderRadius: '8px',
          fontSize: '0.8rem',
          fontWeight: 700,
          cursor: 'pointer'
        }}
      >
        📂 Existant
      </button>
    </div>
  );
};

const EditableTable = ({ fileId, projectId, filename }: { fileId: string; projectId: string; filename?: string }) => {
  const storageKey = `tab-data-${projectId}-${fileId}`;

  const fnL = filename?.toLowerCase() || '';
  const isRACI = fnL.includes('raci');
  const isRisque = fnL.includes('risque');
  const isBudget = fnL.includes('budget');

  const [data, setData] = useState<string[][]>(() => {
    const saved = localStorage.getItem(storageKey);
    if (saved) {
      try {
        return JSON.parse(saved);
      } catch (e) {
        console.error("Error parsing local storage data for", storageKey, e);
      }
    }

    if (isRACI) {
      return [
        ['PH-01', 'Initialisation', 'Définir la charte', 'A', 'R', 'C', 'C', ''],
        ['', '', 'Valider budget', 'I', 'A', 'R', 'I', ''],
        ['PH-02', 'Planification', 'Rédiger WBS', 'I', 'A', 'R', 'C', ''],
        ['', '', 'Répartir tâches', 'C', 'A', 'R', 'I', ''],
        ['PH-03', 'Exécution', 'Développer macro', 'I', 'A', 'I', 'R', '']
      ];
    } else if (isRisque) {
      return [
        ['RSQ-01', 'Départ technique clé', 'Forte', 'Critique', 'À faire', 'Contingence'],
        ['RSQ-02', 'Retard fournisseur', 'Moyenne', 'Basse', 'En cours', 'Prévention'],
        ['RSQ-03', 'Dépassement budget', 'Bas', 'Normale', 'Terminé', 'Surveillance']
      ];
    } else if (isBudget) {
      return [
        ['Logiciels', '01/03/2026', '1 500 000', '1 400 000', '-100 000', 'Payé'],
        ['Consultants', '15/04/2026', '5 000 000', '5 000 000', '0', 'En attente'],
        ['Matériels', '10/05/2026', '2 000 000', '', '', 'À planifier']
      ];
    }
    return [
      ['Taux d\'avancement', 'Responsable A', 'En cours', 'Haute', '15 000 000', '01/04/2026', '15/05/2026', 'Phase de cadrage validée'],
      ['Livrables validés', 'Responsable B', 'Terminé', 'Normale', '2 500 000', '15/03/2026', '30/03/2026', 'Document signé'],
      ['Risques critiques', 'Responsable C', 'À faire', 'Critique', '1 200 000', '01/06/2026', '30/08/2026', 'Attente budget']
    ];
  });

  const save = (newData: string[][]) => {
    setData(newData);
    localStorage.setItem(storageKey, JSON.stringify(newData));
  };

  const handleChange = (r: number, c: number, val: string) => {
    let newVal = val;
    // Macro RACI Validation rules applied!
    if (isRACI && c >= 3 && c <= 6) {
      newVal = val.toUpperCase().replace(/[^RACI\s]/g, '');
      newVal = Array.from(new Set(newVal.replace(/\s/g, '').split(''))).join(' ');
    }
    const newData = [...data];
    newData[r][c] = newVal;
    save(newData);
  };

  const addRow = () => {
    save([...data, Array(8).fill('')]);
  };

  const deleteRow = (rIdx: number) => {
    const newData = data.filter((_, idx) => idx !== rIdx);
    save(newData);
  };

  const clearTable = () => {
    save(Array.from({ length: 5 }, () => Array(8).fill('')));
  };

  let columns = ['Indicateur', 'Acteur', 'Statut', 'Priorité', 'Coût estimé', 'Date Début', 'Date Fin', 'Observations'];
  if (isRACI) columns = ['Code Phase', 'Phase', 'Tâche/Livrable', 'Sponsor', 'Chef Projet', 'Tech Lead', 'Dév.', 'Observ.'];
  else if (isRisque) columns = ['Code', 'Indicateur de Risque', 'Probabilité', 'Impact', 'Statut', 'Action Prévue'];
  else if (isBudget) columns = ['Designation', 'Date', 'Prévu (XOF)', 'Réel (XOF)', 'Écart', 'Statut Paiement'];

  const getCellStyles = (cIdx: number, val: string) => {
    if (!isRACI || cIdx < 3 || cIdx > 6) return {};
    let color = '';
    let fontWeight = 400;
    let bg = 'transparent';
    if (val.includes('R')) { color = '#fff'; bg = '#ef4444'; fontWeight = 800; }
    else if (val.includes('A')) { color = '#fff'; bg = '#3b82f6'; fontWeight = 800; }
    else if (val.includes('C')) { color = '#fff'; bg = '#f59e0b'; fontWeight = 800; }
    else if (val.includes('I')) { color = '#fff'; bg = '#10b981'; fontWeight = 800; }
    return { color, fontWeight, backgroundColor: bg, textAlign: 'center' as const, borderRadius: '6px' };
  };

  return (
    <div className="dynamic-table-container">
      <div className="dynamic-table-toolbar">
        <button className="btn-action macro" onClick={addRow}>+ Ajouter Ligne (Macro)</button>
        <button className="btn-action macro" onClick={clearTable}>🗑️ Vider Tableau</button>
        <button className="btn-action" onClick={() => {
          if (!data || data.length === 0) return;
          const rowsObj = data.map((row: string[]) => {
            const obj: any = {};
            columns.forEach((col, idx) => {
              obj[col] = row[idx] || '';
            });
            return obj;
          });
          exportToCSV(rowsObj, 'export_tableau.csv');
        }}>📥 Exporter Excel</button>
      </div>
      <div className="dynamic-table">
        <div className="table-row header">
          <div className="table-cell row-idx">#</div>
          {columns.map(c => <div key={c} className="table-cell col-header">{c}</div>)}
          <div className="table-cell row-action">Actions</div>
        </div>
        {data.map((row: string[], rIdx: number) => (
          <div key={`row-${rIdx}`} className="table-row">
            <div className="table-cell row-idx">{rIdx + 1}</div>
            {row.map((cell: string, cIdx: number) => {
              const colName = columns[cIdx];
              if (colName === 'Statut' || colName === 'Statut Paiement') {
                return (
                  <select key={`cell-${rIdx}-${cIdx}`} className="table-dropdown" value={cell} onChange={(e) => handleChange(rIdx, cIdx, e.target.value)}>
                    <option value="">Sélectionner...</option>
                    <option value="À faire">À faire / En attente</option>
                    <option value="En cours">En cours</option>
                    <option value="Terminé">Terminé / Payé</option>
                  </select>
                )
              }
              if (colName === 'Priorité' || colName === 'Impact') {
                return (
                  <select key={`cell-${rIdx}-${cIdx}`} className="table-dropdown" value={cell} onChange={(e) => handleChange(rIdx, cIdx, e.target.value)}>
                    <option value="">Sélectionner...</option>
                    <option value="Basse">Basse</option>
                    <option value="Normale">Normale</option>
                    <option value="Haute">Haute</option>
                    <option value="Critique">Critique</option>
                  </select>
                )
              }
              if (colName === 'Probabilité') {
                return (
                  <select key={`cell-${rIdx}-${cIdx}`} className="table-dropdown" value={cell} onChange={(e) => handleChange(rIdx, cIdx, e.target.value)}>
                    <option value="">Sélectionner...</option>
                    <option value="Faible">Faible</option>
                    <option value="Moyenne">Moyenne</option>
                    <option value="Forte">Forte</option>
                  </select>
                )
              }
              return (
                <input
                  key={`cell-${rIdx}-${cIdx}`}
                  className="table-cell-input"
                  style={getCellStyles(cIdx, cell)}
                  value={cell}
                  onChange={(e) => handleChange(rIdx, cIdx, e.target.value)}
                  placeholder="..."
                  spellCheck="false"
                />
              )
            })}
            <div className="table-cell row-action">
              <button className="btn-delete-row" title="Supprimer la ligne" onClick={() => deleteRow(rIdx)}>❌</button>
            </div>
          </div>
        ))}
      </div>
    </div>
  );
};


// ------------- GANTT DATA & VIEWS -------------
const PLANNING_RESOURCES = [
  { id: 1, name: 'Chantal Dupont', role: 'Chef de Projet', email: 'DupontC@confassis.com', tel: '' },
  { id: 2, name: 'Olivier Robespierre', role: 'Sponsor', email: 'RobespierreO@confassis.com', tel: '' },
  { id: 3, name: 'Jacques Meunier', role: 'Responsable Intégration', email: 'MeunierJ@confassis.com', tel: '' },
  { id: 4, name: 'Monique Davant', role: 'Coordinatrice Finances', email: 'DavantM@confassis.com', tel: '' },
  { id: 5, name: 'Denise Attalie', role: 'Coordinatrice Ventes', email: 'AttalieD@confassis.com', tel: '' },
  { id: 6, name: 'Jean Albert', role: 'Coordinateur Achats', email: 'AlbertJ@confassis.com', tel: '' },
  { id: 7, name: 'Eric Duteil', role: 'Coordinateur Logistique', email: 'DuteilE@confassis.com', tel: '' },
  { id: 8, name: 'Daniel Château', role: 'Directeur Informatique', email: 'ChateauD@confassis.com', tel: '' },
  { id: 9, name: 'Bernard Dupuis', role: 'Responsable du Changement', email: 'DupuisB@confassis.com', tel: '' },
  { id: 10, name: 'Marianne Anselme', role: 'Expert SAP Finances FI/CO', email: 'AnselmeM@confassis.com', tel: '' },
  { id: 11, name: 'Gérard Lourte', role: 'Expert SAP Finances Immobilisations', email: 'LourteG@confassis.com', tel: '' },
  { id: 12, name: 'Marcel Poste', role: 'Expert SAP Ventes', email: 'PosteM@confassis.com', tel: '' },
  { id: 13, name: 'Bernard Tassin', role: 'Expert SAP Ventes', email: 'TassinB@confassis.com', tel: '' },
  { id: 14, name: 'Edouard Leperre', role: 'Expert SAP Achats', email: 'LeperreE@confassis.com', tel: '' },
  { id: 15, name: 'Carine Refour', role: 'Expert SAP Achats', email: 'RefourC@confassis.com', tel: '' },
  { id: 16, name: 'Henri Saulnier', role: 'Expert SAP Logistique', email: 'SaulnierH@confassis.com', tel: '' },
  { id: 17, name: 'Marc Ellin', role: 'Expert SAP Logistique', email: 'EllinM@confassis.com', tel: '' },
  { id: 18, name: 'Juliette Binot', role: 'Expert Master Data', email: 'BinotJ@confassis.com', tel: '' },
  { id: 19, name: 'Antoine Fleur', role: 'Expert Master Data', email: 'FleurA@confassis.com', tel: '' },
  { id: 20, name: 'Georges Clonet', role: 'Coordinateur Production', email: 'ClonetG@confassis.com', tel: '' },
  { id: 21, name: 'Utilisateurs CMRL', role: 'Testeurs et valideurs', email: '', tel: '' },
];

const PLANNING_GANTT_DATA = [
  { lvl: 1, id: '1', task: 'Initialisation', status: 'TODO', pred: '', owner: 'Chantal Dupont', start: '4/12/21', charge: 48, deadline: '3/29/24', prog: 51, bold: true, type: 'Phase' },
  { lvl: 2, id: '1.1', task: 'Etude d\'opportunité et approbation', status: 'Done', pred: 'Initialisation', owner: 'Chantal Dupont', start: '4/12/21', charge: 48, deadline: '2/28/24', prog: 51, type: 'Tâche' },
  { lvl: 2, id: '1.2', task: 'Préparation de la Charte Projet - Justification', status: 'In progress', pred: 'Initialisation', owner: 'Bernard Dupuis', start: '2/23/24', charge: 4, deadline: '2/28/24', prog: 30, type: 'Tâche' },
  { lvl: 2, id: '1.3', task: 'Estimation coûts/délais fonctionnels Finances', status: 'In progress', pred: 'Etude d\'opportunité et approbation', owner: 'Eric Duteil', start: '3/14/24', charge: 5, deadline: '3/20/24', prog: 50, type: 'Tâche' },
  { lvl: 2, id: '1.4', task: 'Rédaction de la Charte Projet', status: 'Hold', pred: 'Initialisation', owner: 'Olivier Robespierre', start: '3/23/24', charge: 15, deadline: '3/29/24', prog: 15, type: 'Tâche' },
  { lvl: 2, id: '1.5', task: 'Approbation de la Charte Projet', status: 'Done', pred: 'Initialisation', owner: 'Eric Duteil', start: '4/12/21', charge: 20, deadline: '5/7/21', prog: 100, type: 'Tâche' },
  { lvl: 1, id: '2', task: 'Planification', status: 'TODO', pred: 'Initialisation', owner: 'Denise Attalie', start: '4/19/23', charge: 237, deadline: '8/14/23', prog: 58, bold: true, type: 'Phase' },
  { lvl: 2, id: '2.1', task: 'Conception générale', status: 'TODO', pred: 'Planification', owner: 'Monique Davant', start: '4/22/23', charge: 35, deadline: '6/15/23', prog: 50, type: 'Tâche' },
  { lvl: 2, id: '2.2', task: 'Rédaction du cahier des charges', status: 'Done', pred: 'Approbation de la Charte Projet', owner: 'Marcel Poste', start: '4/22/23', charge: 50, deadline: '7/6/23', prog: 90, type: 'Tâche' },
  { lvl: 2, id: '2.3', task: 'Etablissement du Planning', status: 'In progress', pred: 'Conception générale', owner: 'Daniel Château', start: '5/19/23', charge: 60, deadline: '8/14/23', prog: 50, type: 'Tâche' },
  { lvl: 2, id: '2.4', task: 'Etablissement Plan de Communication', status: 'In progress', pred: 'Conception générale', owner: 'Denise Attali', start: '4/19/23', charge: 33, deadline: '6/8/23', prog: 50, type: 'Tâche' },
  { lvl: 2, id: '2.5', task: 'Approbation Cahier des charges, Plan de projet et Planning', status: 'In Hold', pred: 'Conception générale', owner: 'Marcel Poste', start: '5/15/23', charge: 59, deadline: '8/8/23', prog: 50, type: 'Tâche' },
  { lvl: 1, id: '3', task: 'Réalisation', status: 'TODO', pred: 'Rédaction du cahier des charges', owner: 'Georges Clonet', start: '6/11/23', charge: 504, deadline: '2/22/24', prog: 31, bold: true, type: 'Phase' },
  { lvl: 2, id: '3.1', task: 'Paramétrage et tests unitaires', status: 'Done', pred: '', owner: 'Denise Attali', start: '6/11/23', charge: 14, deadline: '6/29/23', prog: 20, type: 'Tâche' },
  { lvl: 2, id: '3.2', task: 'Paramétrage et corrections Master Data', status: 'In progress', pred: 'Planification', owner: 'Georges Clonet', start: '7/13/23', charge: 121, deadline: '1/3/24', prog: 13, type: 'Tâche' },
  { lvl: 2, id: '3.3', task: 'Paramétrage, programmation et test Production', status: 'TODO', pred: 'Planification', owner: 'Bernard Dupuis', start: '8/13/23', charge: 80, deadline: '12/19/23', prog: 15, type: 'Tâche' },
  { lvl: 2, id: '3.4', task: 'Tests, corrections et validation du système', status: 'In progress', pred: 'Planification', owner: 'Carine Refour', start: '9/13/23', charge: 59, deadline: '12/5/23', prog: 45, type: 'Tâche' },
  { lvl: 2, id: '3.5', task: 'Validation officielle par les utilisateurs', status: 'TODO', pred: 'Paramétrage et tests unitaires', owner: 'Antoine Fleur', start: '6/11/23', charge: 180, deadline: '2/22/24', prog: 60, type: 'Tâche' },
  { lvl: 2, id: '3.6', task: 'Validation de la phase par le comité de direction', status: 'Hold', pred: 'Tests, corrections et validation du système', owner: 'Henri Saulnier', start: '9/29/23', charge: 40, deadline: '11/24/23', prog: 30, type: 'Tâche' },
  { lvl: 1, id: '4', task: 'Contrôle et suivi', status: 'TODO', pred: 'Validation officielle par les utilisateurs', owner: 'Edouard Leperre', start: '1/18/23', charge: 579, deadline: '5/15/25', prog: 32, bold: true, type: 'Phase' },
  { lvl: 2, id: '4.1', task: 'Tests d\'intégration et de consolidation', status: 'Done', pred: '', owner: 'Jean Albert', start: '12/8/23', charge: 91, deadline: '4/15/24', prog: 20, type: 'Tâche' },
  { lvl: 2, id: '4.2', task: 'Tests d\'intégration et consolidation, validation du système', status: 'Done', pred: 'Réalisation', owner: 'Bernard Dupuis', start: '4/8/25', charge: 28, deadline: '5/15/25', prog: 60, type: 'Tâche' },
  { lvl: 1, id: '5', task: 'Clôture', status: 'TODO', pred: 'Tests d\'intégration et consolidation', owner: 'Denise Attali', start: '7/2/23', charge: 126, deadline: '10/4/23', prog: 36, bold: true, type: 'Phase' },
];

const getStatusColor = (status: string) => {
  if (status === 'Done') return '#10b981';
  if (status === 'In progress') return '#3b82f6';
  return '#94a3b8';
};


const GanttTableView = ({ rows, onChange, onAddTask, onAddMilestone, onRemove, onReorder }: { rows: any[], onChange: any, onAddTask: any, onAddMilestone: any, onRemove: any, onReorder: (draggedId: string, targetId: string) => void }) => {
  const [draggedId, setDraggedId] = useState<string | null>(null);

  const handleDragStart = (e: React.DragEvent, id: string) => {
    setDraggedId(id);
    e.dataTransfer.setData('text/plain', id);
    e.dataTransfer.effectAllowed = 'move';
  };

  const handleDragOver = (e: React.DragEvent) => {
    e.preventDefault();
    e.dataTransfer.dropEffect = 'move';
  };

  const handleDrop = (e: React.DragEvent, targetId: string) => {
    e.preventDefault();
    const id = e.dataTransfer.getData('text/plain');
    if (id && id !== targetId) {
      onReorder(id, targetId);
    }
    setDraggedId(null);
  };
  return (
    <div style={{ flex: 1, backgroundColor: 'var(--main-bg)', color: 'var(--main-text)', padding: '1.5rem', display: 'flex', flexDirection: 'column', overflow: 'hidden' }}>
      <datalist id="task-list">
        {rows.map((r: any) => <option key={r.id} value={r.task} />)}
      </datalist>
      <div style={{ marginBottom: '1rem', display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
        <h2 style={{ margin: 0, color: 'var(--table-primary)', display: 'flex', alignItems: 'center', gap: '0.5rem', fontWeight: 800 }}>
          📊 Données du planning
        </h2>
      </div>

      <div style={{ background: 'rgba(59, 130, 246, 0.1)', borderLeft: '4px solid #3b82f6', padding: '12px 15px', color: '#2563eb', fontWeight: 'bold', fontSize: '0.85rem', marginBottom: '1.5rem', borderRadius: '0 8px 8px 0' }}>
        🚀 Vision tabulaire dynamique du WBS et des échéances (Précisez Jalons, Phases et Tâches).
      </div>

      <div style={{ flex: 1, overflow: 'auto', background: 'var(--main-surface)', borderRadius: '12px', border: '1px solid var(--main-border)', boxShadow: '0 4px 6px -1px rgba(0,0,0,0.05)' }}>
        <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '0.8rem', minWidth: '1200px' }}>
          <thead style={{ background: 'var(--table-primary)', color: '#fff', position: 'sticky', top: 0, zIndex: 10 }}>
            <tr>
              <th style={{ padding: '12px', border: '1px solid rgba(255,255,255,0.1)' }}>ID</th>
              <th style={{ padding: '12px', border: '1px solid rgba(255,255,255,0.1)', width: '150px' }}>Type</th>
              <th style={{ padding: '12px', border: '1px solid rgba(255,255,255,0.1)', textAlign: 'left', width: '450px' }}>WBS - Activité / tâche</th>
              <th style={{ padding: '12px', border: '1px solid rgba(255,255,255,0.1)' }}>Statut</th>
              <th style={{ padding: '12px', border: '1px solid rgba(255,255,255,0.1)' }}>Préd.</th>
              <th style={{ padding: '12px', border: '1px solid rgba(255,255,255,0.1)', textAlign: 'left', width: '200px' }}>Attribué à</th>
              <th style={{ padding: '12px', border: '1px solid rgba(255,255,255,0.2)' }}>Début</th>
              <th style={{ padding: '12px', border: '1px solid rgba(255,255,255,0.2)' }}>Charge (J)</th>
              <th style={{ padding: '12px', border: '1px solid rgba(255,255,255,0.2)' }}>Deadline</th>
              <th style={{ padding: '12px', border: '1px solid rgba(255,255,255,0.1)' }}>Progression</th>
              <th style={{ padding: '12px', border: '1px solid rgba(255,255,255,0.1)', width: '120px' }}>Action</th>
            </tr>
          </thead>
          <tbody>
            {rows.map((d: any, i: number) => (
              <tr
                key={i}
                draggable={d.type !== 'Phase'}
                onDragStart={(e) => handleDragStart(e, d.id)}
                onDragOver={handleDragOver}
                onDrop={(e) => handleDrop(e, d.id)}
                style={{
                  borderBottom: '1px solid var(--main-border)',
                  background: d.type === 'Phase' ? '#059669' : d.bold ? 'rgba(59, 130, 246, 0.05)' : i % 2 === 0 ? 'transparent' : 'rgba(0,0,0,0.01)',
                  color: d.type === 'Phase' ? '#ffffff' : 'inherit',
                  fontWeight: d.type === 'Phase' || d.bold ? 800 : 400,
                  opacity: draggedId === d.id ? 0.5 : 1,
                  cursor: d.type !== 'Phase' ? 'grab' : 'default',
                  boxShadow: d.type === 'Phase' ? 'inset 0 2px 4px rgba(0,0,0,0.1)' : 'none'
                }}
              >

                <td style={{
                  padding: '10px',
                  textAlign: 'center',
                  color: d.type === 'Phase' ? '#fff' : (d.type === 'Jalon' ? '#f59e0b' : 'var(--table-primary)'),
                  fontWeight: 900
                }}>{d.id}</td>
                <td style={{ padding: '10px' }}>
                  <select
                    value={d.type || (d.bold ? (d.charge === 0 ? 'Jalon' : 'Phase') : 'Tâche')}
                    onChange={(e) => onChange(d.id, 'type', e.target.value)}
                    style={{ background: d.type === 'Phase' ? 'transparent' : 'rgba(255,255,255,0.05)', color: d.type === 'Phase' ? '#fff' : 'inherit', border: d.type === 'Phase' ? '1px solid rgba(255,255,255,0.4)' : '1px solid var(--main-border)', borderRadius: '4px', fontSize: '0.75rem', padding: '2px 4px', width: '100%', fontWeight: d.type === 'Phase' ? 700 : 400 }}
                    disabled={d.type === 'Phase'}
                  >
                    <option value="Phase">Phase</option>
                    <option value="Tâche">Tâche</option>
                    <option value="Jalon">Jalon</option>
                  </select>
                </td>
                <td style={{ padding: '10px', paddingLeft: d.lvl > 1 ? `${d.lvl * 20}px` : '15px' }}>
                  <textarea
                    value={d.task}
                    onChange={(e) => onChange(d.id, 'task', e.target.value)}
                    rows={1}
                    style={{
                      width: '100%',
                      border: 'none',
                      background: 'transparent',
                      color: d.type === 'Jalon' ? '#b45309' : 'inherit',
                      fontWeight: d.type === 'Phase' || d.type === 'Jalon' || d.bold ? 800 : 400,
                      resize: 'none',
                      fontFamily: 'inherit',
                      fontSize: 'inherit',
                      overflow: 'hidden',
                      display: 'block'
                    }}
                    disabled={d.type === 'Phase'}
                    onInput={(e: any) => {
                      e.target.style.height = 'auto';
                      e.target.style.height = e.target.scrollHeight + 'px';
                    }}
                    ref={(el) => {
                      if (el) {
                        el.style.height = 'auto';
                        el.style.height = el.scrollHeight + 'px';
                      }
                    }}
                  />
                </td>
                <td style={{ padding: '10px', textAlign: 'center' }}>
                  <select value={d.status} onChange={(e) => onChange(d.id, 'status', e.target.value)} style={{ background: d.type === 'Phase' ? 'rgba(255,255,255,0.2)' : getStatusColor(d.status) + '22', color: d.type === 'Phase' ? '#fff' : getStatusColor(d.status), padding: '4px 10px', borderRadius: '12px', fontSize: '0.7rem', fontWeight: 800, border: 'none', cursor: 'pointer' }}>
                    <option value="TODO">TODO</option>
                    <option value="In progress">In progress</option>
                    <option value="Done">Done</option>
                    <option value="Hold">Hold</option>
                  </select>
                </td>
                <td style={{ padding: '10px', textAlign: 'center' }}>
                  <input list="task-list" value={d.pred} onChange={(e) => onChange(d.id, 'pred', e.target.value)} style={{ width: '120px', border: 'none', background: 'transparent', textAlign: 'center', color: 'inherit', fontSize: '0.75rem' }} placeholder="Précédent..." />
                </td>
                <td style={{ padding: '10px' }}>
                  <input value={d.owner} onChange={(e) => onChange(d.id, 'owner', e.target.value)} style={{ width: '100%', border: 'none', background: 'transparent', color: 'inherit' }} />
                </td>
                <td style={{ padding: '10px', textAlign: 'center' }}>
                  <input value={d.start} onChange={(e) => onChange(d.id, 'start', e.target.value)} style={{ width: '80px', border: 'none', background: 'transparent', textAlign: 'center', color: 'inherit' }} />
                </td>
                <td style={{ padding: '10px', textAlign: 'center' }}>
                  <input type="number" value={d.charge} onChange={(e) => onChange(d.id, 'charge', parseInt(e.target.value))} style={{ width: '40px', border: 'none', background: 'transparent', textAlign: 'center', color: 'inherit' }} />
                </td>
                <td style={{ padding: '10px', textAlign: 'center' }}>
                  <input value={d.deadline} onChange={(e) => onChange(d.id, 'deadline', e.target.value)} style={{ width: '80px', border: 'none', background: 'transparent', textAlign: 'center', color: 'inherit' }} />
                </td>
                <td style={{ padding: '10px', textAlign: 'center' }}>
                  <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
                    <input type="number" value={d.prog} onChange={(e) => onChange(d.id, 'prog', parseInt(e.target.value))} style={{ width: '40px', border: 'none', background: 'transparent', textAlign: 'right', color: 'inherit', fontSize: '0.8rem' }} />
                    <span style={{ fontSize: '0.7rem' }}>%</span>
                  </div>
                </td>
                <td style={{ padding: '10px', textAlign: 'center' }}>
                  <div style={{ display: 'flex', gap: '4px', justifyContent: 'center' }}>
                    {d.type !== 'Phase' && (
                      <button onClick={() => onRemove(d.id)} style={{ background: 'none', border: 'none', cursor: 'pointer' }} title="Supprimer">🗑️</button>

                    )}
                  </div>
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
      <div style={{ marginTop: '1rem', display: 'flex', gap: '1rem', flexWrap: 'wrap' }}>
        <button onClick={onAddTask} style={{ padding: '10px 20px', background: 'var(--table-primary)', color: '#fff', border: 'none', borderRadius: '8px', cursor: 'pointer', fontWeight: 800 }}>
          + Ajouter une tâche
        </button>
        <button onClick={onAddMilestone} style={{ padding: '10px 20px', background: '#f59e0b', color: '#fff', border: 'none', borderRadius: '8px', cursor: 'pointer', fontWeight: 800 }}>
          + Ajouter jalon
        </button>
      </div>
    </div>
  );
};

const MacroPlanningView = ({ fileId, projectId, currentDossierId, initialView }: { fileId: string; projectId: string; currentDossierId: string | null; initialView?: string }) => {
  const [viewMode, setViewMode] = useState(initialView || 'option1');
  const dossierSubKey = currentDossierId ? `-${currentDossierId}` : "";
  const storageKey = `faab-macro-planning-data-${projectId}-${fileId}${dossierSubKey}`;
  const resourceKey = `faab-macro-planning-resources-${projectId}-${fileId}${dossierSubKey}`;

  const [rows, setRows] = useState(() => {
    const saved = localStorage.getItem(storageKey);
    return saved ? JSON.parse(saved) : PLANNING_GANTT_DATA;
  });

  const [resources, setResources] = useState(() => {
    const saved = localStorage.getItem(resourceKey);
    return saved ? JSON.parse(saved) : PLANNING_RESOURCES;
  });

  useEffect(() => {
    localStorage.setItem(storageKey, JSON.stringify(rows));
  }, [rows, storageKey]);

  useEffect(() => {
    localStorage.setItem(resourceKey, JSON.stringify(resources));
  }, [resources, resourceKey]);

  useEffect(() => {
    if (initialView) setViewMode(initialView);
  }, [initialView]);

  const handleRowChange = (id: string, field: string, val: any) => {
    setRows(rows.map((r: any) => {
      if (r.id !== id) return r;
      let updatedRow = { ...r, [field]: val };

      const parseDMY = (s: string) => {
        if (!s) return null;
        const parts = s.split('/');
        if (parts.length !== 3) return null;
        let d = parseInt(parts[0]);
        let m = parseInt(parts[1]) - 1;
        let y = parseInt(parts[2]);
        if (isNaN(d) || isNaN(m) || isNaN(y)) return null;
        if (y < 100) y += 2000;
        return new Date(y, m, d);
      };

      const formatDMY = (d: Date) => {
        return `${d.getDate().toString().padStart(2, '0')}/${(d.getMonth() + 1).toString().padStart(2, '0')}/${d.getFullYear()}`;
      };

      if (field === 'start' || field === 'charge') {
        const startDate = parseDMY(updatedRow.start);
        const charge = parseInt(updatedRow.charge);
        if (startDate && !isNaN(charge)) {
          const endDate = new Date(startDate);
          endDate.setDate(startDate.getDate() + charge);
          updatedRow.deadline = formatDMY(endDate);
        }
      } else if (field === 'deadline') {
        const startDate = parseDMY(updatedRow.start);
        const endDate = parseDMY(updatedRow.deadline);
        if (startDate && endDate) {
          const diffTime = endDate.getTime() - startDate.getTime();
          const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));
          updatedRow.charge = diffDays >= 0 ? diffDays : 0;
        }
      }

      return updatedRow;
    }));
  };

  const handleResourceChange = (id: any, field: string, val: any) => {
    setResources(resources.map((r: any) => r.id === id ? { ...r, [field]: val } : r));
  };



  const onAddTask = () => {
    const phases = rows.filter((r: any) => r.type === 'Phase');
    if (phases.length === 0) { alert("Créez au moins une Phase d'abord dans Paramétrage (ou contactez l'admin)."); return; }
    const phaseId = window.prompt("Dans quelle Phase (ID) voulez-vous ajouter cette tâche ?\n" + phases.map((l: any) => `${l.id}: ${l.task}`).join("\n"));
    const phase = phases.find((l: any) => l.id === phaseId);
    if (!phase) return;

    const phaseIndex = rows.findIndex((r: any) => r.id === phase.id);
    let insertIndex = phaseIndex + 1;
    while (insertIndex < rows.length && rows[insertIndex].type !== 'Phase') { insertIndex++; }

    const today = new Date().toLocaleDateString('fr-FR');
    const newRow = { lvl: 2, id: `${phase.id}.${insertIndex - phaseIndex}`, task: 'Nouvelle tâche', status: 'TODO', pred: '', owner: '', start: today, charge: 5, deadline: today, prog: 0, type: 'Tâche', bold: false };
    const newRows = [...rows];
    newRows.splice(insertIndex, 0, newRow);
    setRows(newRows);
  };

  const onAddMilestone = () => {
    const phases = rows.filter((r: any) => r.type === 'Phase');
    if (phases.length === 0) { alert("Créez au moins une Phase d'abord."); return; }
    const phaseId = window.prompt("Dans quelle Phase (ID) voulez-vous ajouter ce jalon ?\n" + phases.map((l: any) => `${l.id}: ${l.task}`).join("\n"));
    const phase = phases.find((l: any) => l.id === phaseId);
    if (!phase) return;

    const phaseIndex = rows.findIndex((r: any) => r.id === phase.id);
    let insertIndex = phaseIndex + 1;
    while (insertIndex < rows.length && rows[insertIndex].type !== 'Phase') { insertIndex++; }

    const today = new Date().toLocaleDateString('fr-FR');
    const newRow = { lvl: 2, id: `${phase.id}.${insertIndex - phaseIndex}`, task: 'Nouveau Jalon', status: 'TODO', pred: '', owner: '', start: today, charge: 0, deadline: today, prog: 0, type: 'Jalon', bold: true };
    const newRows = [...rows];
    newRows.splice(insertIndex, 0, newRow);
    setRows(newRows);
  };



  const onReorder = (draggedId: string, targetId: string) => {
    const dragIdx = rows.findIndex((r: any) => r.id === draggedId);
    const targetIdx = rows.findIndex((r: any) => r.id === targetId);
    if (dragIdx === -1 || targetIdx === -1 || dragIdx === targetIdx) return;

    const draggedRow = rows[dragIdx];
    const targetRow = rows[targetIdx];

    // Cannot drag Phases
    if (draggedRow.type === 'Phase') return;

    // Cannot drop on a Phase if it means crossing boundaries
    // We check if the dragged item and the item at target index belong to the same phase section
    // A simple way is to find the previous Phase for both
    const findPhaseId = (index: number) => {
      for (let i = index; i >= 0; i--) {
        if (rows[i].type === 'Phase') return rows[i].id;
      }
      return null;
    };

    const dragPhase = findPhaseId(dragIdx);
    const targetPhase = findPhaseId(targetIdx);

    if (dragPhase !== targetPhase || targetRow.type === 'Phase') {
      // If target is a Phase, we might want to allow dropping just after it if it's the same phase section?
      // But user said "empecher le déplacement des lignes inter-phases".
      // Let's just block if they are not in the same section.
      return;
    }

    const newRows = [...rows];
    const [movedItem] = newRows.splice(dragIdx, 1);
    newRows.splice(targetIdx, 0, movedItem);
    setRows(newRows);
  };

  if (viewMode === 'option2') return <GanttGraphView rows={rows} />;
  if (viewMode === 'option_pert') return <GanttPertView rows={rows} />;
  if (viewMode === 'option3') return <GanttRessourcesView resources={resources} onChange={handleResourceChange} onAdd={() => setResources([...resources, { id: Date.now(), name: 'Nouveau', role: '', email: '', tel: '' }])} onRemove={(id: any) => setResources(resources.filter((r: any) => r.id !== id))} />;
  if (viewMode === 'option4') return <GanttParametrageView />;

  return <GanttTableView
    rows={rows}
    onChange={handleRowChange}
    onAddTask={onAddTask}
    onAddMilestone={onAddMilestone}
    onRemove={(id: string) => {
      const row = rows.find((r: any) => r.id === id);
      if (row && row.type === 'Phase') {
        alert("La suppression des phases est interdite.");
        return;
      }
      setRows(rows.filter((r: any) => r.id !== id));
    }}
    onReorder={onReorder}
  />;
};

const GanttPertView = ({ rows }: { rows: any[] }) => {
  // Calcul dynamique des étapes (layers) pour la disposition en réseau
  const nodes = rows.map((r, i) => ({ ...r, layer: 0, index: i + 1 }));

  let changed = true;
  let iters = 0;
  while (changed && iters < 50) {
    changed = false;
    nodes.forEach(n => {
      if (n.pred) {
        // La tâche a un prédécesseur explicite
        const p = nodes.find(x => x.task === n.pred || x.id === n.pred);
        if (p && p.layer >= n.layer) {
          n.layer = p.layer + 1;
          changed = true;
        }
      } else if (n.type !== 'Phase' && n.index > 1) {
        // La tâche n'a pas de prédécesseur, elle suit la phase en cours
        const prevPhase = nodes.slice(0, n.index - 1).reverse().find(x => x.type === 'Phase');
        if (prevPhase && prevPhase.layer >= n.layer) {
          n.layer = prevPhase.layer + 1;
          changed = true;
        }
      }
    });
    iters++;
  }

  const layers: any[][] = [];
  nodes.forEach(n => {
    if (!layers[n.layer]) layers[n.layer] = [];
    layers[n.layer].push(n);
  });

  return (
    <div style={{ flex: 1, backgroundColor: 'var(--main-bg)', color: 'var(--main-text)', padding: '1.5rem', display: 'flex', flexDirection: 'column', overflow: 'hidden' }}>
      <div style={{ marginBottom: '1rem', display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
        <h2 style={{ margin: 0, color: 'var(--table-primary)', display: 'flex', alignItems: 'center', gap: '0.5rem', fontWeight: 800 }}>
          🕸️ Diagramme de PERT
        </h2>
      </div>

      <div style={{ background: 'rgba(59, 130, 246, 0.1)', borderLeft: '4px solid #3b82f6', padding: '12px 15px', color: '#2563eb', fontWeight: 'bold', fontSize: '0.85rem', marginBottom: '1.5rem', borderRadius: '0 8px 8px 0' }}>
        📊 Visualisation des dépendances (Méthode PERT) structurée en étapes.
      </div>

      <div style={{ flex: 1, overflow: 'auto', background: 'var(--main-surface)', borderRadius: '16px', border: '1px solid var(--main-border)', padding: '2rem', whiteSpace: 'nowrap' }}>

        <div style={{ display: 'inline-flex', flexDirection: 'row', alignItems: 'center', gap: '30px', padding: '2rem' }}>
          {layers.map((layer, lIdx) => (
            <div key={lIdx} style={{ display: 'flex', flexDirection: 'column', gap: '40px', position: 'relative' }}>

              {/* Ligne verticale pointillée si plusieurs tâches en parallèle dans la même étape */}
              {layer.length > 1 && (
                <div style={{ position: 'absolute', left: '19px', top: '20px', bottom: '20px', width: '2px', borderLeft: '2px dashed #3b82f6', zIndex: 0 }}></div>
              )}

              {layer.map((node) => (
                <div key={node.id} style={{ display: 'flex', alignItems: 'center', position: 'relative', zIndex: 1 }}>

                  {/* Cercle d'événement (Nœud Etape) */}
                  <div style={{ width: '40px', height: '40px', borderRadius: '50%', background: '#3b82f6', color: '#fff', display: 'flex', alignItems: 'center', justifyContent: 'center', fontWeight: '900', boxShadow: '0 4px 6px rgba(0,0,0,0.1)' }}>
                    {node.index}
                  </div>

                  {/* Ligne fléchée avec la Box (Activité) */}
                  <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'center', margin: '0 5px', width: '180px' }}>
                    {/* Boîte de la tâche au-dessus de la ligne (Style image référence) */}
                    <div style={{
                      border: node.type === 'Phase' ? 'none' : '2px solid #b45309',
                      background: node.type === 'Phase' ? '#059669' : 'var(--main-surface)',
                      color: node.type === 'Phase' ? '#fff' : 'inherit',
                      padding: '8px 12px',
                      borderRadius: '4px',
                      fontSize: '0.8rem',
                      fontWeight: 'bold',
                      textAlign: 'center',
                      marginBottom: '6px',
                      width: '100%',
                      whiteSpace: 'normal',
                      boxShadow: '0 2px 4px rgba(0,0,0,0.05)',
                      overflow: 'hidden',
                      textOverflow: 'ellipsis'
                    }}>
                      {node.task}
                    </div>

                    {/* Flèche directionnelle */}
                    <div style={{ display: 'flex', alignItems: 'center', width: '100%' }}>
                      <div style={{ height: '2px', background: '#3b82f6', flex: 1 }}></div>
                      <div style={{ width: 0, height: 0, borderTop: '6px solid transparent', borderBottom: '6px solid transparent', borderLeft: '8px solid #3b82f6' }}></div>
                    </div>

                    {/* Durée en dessous de la ligne */}
                    <div style={{ fontSize: '0.85rem', fontWeight: 'bold', marginTop: '6px', opacity: 0.8 }}>
                      {node.type !== 'Phase' ? `${node.charge} j` : ''}
                    </div>
                  </div>

                </div>
              ))}
            </div>
          ))}

          {/* Nœud Final de fin de projet */}
          <div style={{ display: 'flex', alignItems: 'center', marginLeft: '-20px' }}>
            <div style={{ width: '40px', height: '40px', borderRadius: '50%', background: '#10b981', color: '#fff', display: 'flex', alignItems: 'center', justifyContent: 'center', fontWeight: '900', boxShadow: '0 4px 6px rgba(0,0,0,0.1)' }}>
              Fin
            </div>
          </div>

        </div>
      </div>
    </div>
  );
};

const GanttGraphView = ({ rows }: { rows: any[] }) => {
  const [scale, setScale] = useState<'days' | 'months' | 'years'>('years');

  // Helper utils
  const parseDMY = (dateStr: string) => {
    if (!dateStr || dateStr === '-') return null;
    const parts = dateStr.split('/');
    if (parts.length !== 3) return null;
    const d = parseInt(parts[0]);
    const m = parseInt(parts[1]) - 1;
    let y = parseInt(parts[2]);
    if (y < 100) y += 2000;
    return new Date(y, m, d);
  };

  // Find project bounds
  let minDateObj: Date | null = null;
  let maxDateObj: Date | null = null;

  rows.forEach(r => {
    const s = parseDMY(r.start);
    const d = parseDMY(r.deadline);
    if (s && (!minDateObj || s < minDateObj)) minDateObj = s;
    if (d && (!maxDateObj || d > maxDateObj)) maxDateObj = d;
  });

  // Default Fallbacks if no dates
  if (!minDateObj) minDateObj = new Date(2024, 0, 1);
  if (!maxDateObj) maxDateObj = new Date(2025, 11, 31);

  // Buffer margins
  const projStart = new Date(minDateObj.getFullYear(), minDateObj.getMonth(), 1);
  const projEnd = new Date(maxDateObj.getFullYear(), maxDateObj.getMonth() + 1, 0);

  // Dynamic header generation
  const getHeaders = () => {
    const headers: { label: string; width: number; key: string }[] = [];
    const curr = new Date(projStart);

    if (scale === 'years') {
      while (curr <= projEnd) {
        const year = curr.getFullYear();
        headers.push({ label: `${year}`, width: 1, key: `y-${year}` });
        curr.setFullYear(year + 1);
      }
    } else if (scale === 'months') {
      const monthNames = ["Jan", "Fév", "Mar", "Avr", "Mai", "Juin", "Juil", "Août", "Sep", "Oct", "Nov", "Déc"];
      while (curr <= projEnd) {
        const month = curr.getMonth();
        const year = curr.getFullYear();
        headers.push({ label: `${monthNames[month]} ${year}`, width: 1, key: `m-${year}-${month}` });
        curr.setMonth(month + 1);
      }
    } else { // Days
      while (curr <= projEnd) {
        const day = curr.getDate();
        const month = curr.getMonth() + 1;
        headers.push({ label: `${day}/${month}`, width: 1, key: `d-${curr.getTime()}` });
        curr.setDate(day + 1);
      }
    }
    return headers;
  };

  const timeHeaders = getHeaders();
  const totalSlots = timeHeaders.length;

  const getPosition = (dateStr: string) => {
    const d = parseDMY(dateStr);
    if (!d) return 0;
    const startTime = projStart.getTime();
    const endTime = projEnd.getTime();
    const p = ((d.getTime() - startTime) / (endTime - startTime)) * 100;
    return Math.max(0, Math.min(100, p));
  };

  return (
    <div style={{ flex: 1, backgroundColor: 'var(--main-bg)', color: 'var(--main-text)', padding: '1.5rem', display: 'flex', flexDirection: 'column', overflow: 'hidden' }}>
      <div style={{ marginBottom: '1.5rem', display: 'flex', justifyContent: 'space-between', alignItems: 'center', flexWrap: 'wrap', gap: '1rem' }}>
        <h2 style={{ margin: 0, color: 'var(--table-primary)', display: 'flex', alignItems: 'center', gap: '0.5rem', fontWeight: 800 }}>
          📅 Graphique de Gantt
        </h2>

        <div style={{ display: 'flex', background: 'var(--main-surface)', padding: '4px', borderRadius: '10px', border: '1px solid var(--main-border)', boxShadow: '0 2px 5px rgba(0,0,0,0.05)' }}>
          <button
            onClick={() => setScale('years')}
            style={{ padding: '6px 16px', border: 'none', borderRadius: '8px', cursor: 'pointer', fontSize: '0.8rem', fontWeight: 700, background: scale === 'years' ? 'var(--table-primary)' : 'transparent', color: scale === 'years' ? '#fff' : 'var(--main-text)', transition: 'all 0.2s' }}>
            Années
          </button>
          <button
            onClick={() => setScale('months')}
            style={{ padding: '6px 16px', border: 'none', borderRadius: '8px', cursor: 'pointer', fontSize: '0.8rem', fontWeight: 700, background: scale === 'months' ? 'var(--table-primary)' : 'transparent', color: scale === 'months' ? '#fff' : 'var(--main-text)', transition: 'all 0.2s' }}>
            Mois
          </button>
          <button
            onClick={() => setScale('days')}
            style={{ padding: '6px 16px', border: 'none', borderRadius: '8px', cursor: 'pointer', fontSize: '0.8rem', fontWeight: 700, background: scale === 'days' ? 'var(--table-primary)' : 'transparent', color: scale === 'days' ? '#fff' : 'var(--main-text)', transition: 'all 0.2s' }}>
            Jours
          </button>
        </div>
      </div>

      <div style={{ flex: 1, overflow: 'auto', border: '1px solid var(--main-border)', borderRadius: '12px', display: 'flex', flexDirection: 'column', background: 'var(--main-surface)', boxShadow: '0 4px 20px rgba(0,0,0,0.05)' }}>
        {/* Header timeline */}
        <div style={{ display: 'flex', background: 'var(--table-primary)', color: '#fff', fontWeight: 900, fontSize: '0.75rem', position: 'sticky', top: 0, zIndex: 10 }}>
          <div style={{ width: '300px', minWidth: '300px', padding: '15px', borderRight: '1px solid rgba(255,255,255,0.2)', backgroundColor: 'var(--table-primary)' }}>Tâches du projet</div>
          <div style={{ flex: 1, display: 'flex' }}>
            {timeHeaders.map(h => (
              <div key={h.key} style={{ flex: 1, textAlign: 'center', padding: '12px 5px', borderRight: '1px solid rgba(255,255,255,0.1)', minWidth: scale === 'days' ? '120px' : scale === 'months' ? '100px' : '150px' }}>{h.label}</div>
            ))}
          </div>
        </div>

        {/* Rows with Bars */}
        <div style={{ flex: 1, overflowY: 'visible', position: 'relative', display: 'flex', flexDirection: 'column' }}>
          {/* Grid Lines */}
          <div style={{ position: 'absolute', top: 0, left: '300px', right: 0, bottom: 0, display: 'flex', pointerEvents: 'none' }}>
            {timeHeaders.map(h => (
              <div key={`grid-${h.key}`} style={{ flex: 1, borderRight: '1px solid var(--main-border)', opacity: 0.2, minWidth: scale === 'days' ? '120px' : scale === 'months' ? '100px' : '150px' }} />
            ))}
          </div>

          {rows.map((d: any, i: number) => {
            const startP = getPosition(d.start);
            const endP = getPosition(d.deadline);
            const width = Math.max(0.2, endP - startP);

            return (
              <div key={i} style={{ display: 'flex', borderBottom: '1px solid var(--main-border)', background: d.type === 'Phase' ? 'rgba(5, 150, 105, 0.05)' : 'transparent', position: 'relative', zIndex: 1, minHeight: '45px' }}>
                <div style={{
                  width: '300px', minWidth: '300px', padding: '12px 15px', paddingLeft: d.lvl > 1 ? `${d.lvl * 20}px` : '15px',
                  fontWeight: d.type === 'Phase' || d.type === 'Jalon' || d.bold ? 800 : 500, fontSize: '0.8rem', borderRight: '1px solid var(--main-border)',
                  whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis', color: d.type === 'Phase' ? '#059669' : (d.type === 'Jalon' ? '#f59e0b' : 'var(--main-text)'),
                  position: 'sticky', left: 0, zIndex: 5, backgroundColor: 'var(--main-surface)'
                }}>
                  <span style={{ color: d.type === 'Phase' ? '#059669' : (d.type === 'Jalon' ? '#f59e0b' : 'var(--table-primary)'), marginRight: '8px', opacity: 0.8 }}>{d.id}</span> {d.task}
                </div>

                <div style={{ flex: 1, position: 'relative', padding: '10px 0', minWidth: `${totalSlots * (scale === 'days' ? 120 : scale === 'months' ? 100 : 150)}px` }}>
                  {d.type === 'Jalon' ? (
                    /* Rendu pour les Jalons (Losange) */
                    <div style={{
                      position: 'absolute',
                      left: `${startP}%`,
                      transform: 'translateX(-50%) rotate(45deg)',
                      width: '18px',
                      height: '18px',
                      background: '#f59e0b', /* Couleur orange/ambre distinctive pour les jalons */
                      border: '2px solid #fff',
                      boxShadow: '0 2px 5px rgba(0,0,0,0.2)',
                      zIndex: 3,
                      cursor: 'pointer',
                      top: '13px'
                    }} title={`Jalon: ${d.task} (${d.start})`} />
                  ) : (
                    /* Rendu pour les Tâches et Phases (Barres) */
                    <div style={{
                      position: 'absolute',
                      left: `${startP}%`,
                      width: `${width}%`,
                      height: d.type === 'Phase' ? '26px' : '22px',
                      background: d.type === 'Phase' ? '#059669' : getStatusColor(d.status),
                      borderRadius: '6px',
                      display: 'flex',
                      alignItems: 'center',
                      color: '#fff',
                      fontSize: '0.7rem',
                      fontWeight: 900,
                      padding: '0 10px',
                      overflow: 'hidden',
                      boxShadow: '0 4px 8px rgba(0,0,0,0.1)',
                      transition: 'transform 0.2s',
                      cursor: 'pointer',
                      zIndex: 2
                    }} title={`${d.task} (${d.start} - ${d.deadline})`}>
                      <span style={{ whiteSpace: 'nowrap' }}>{width > 3 ? `${d.prog}%` : ''}</span>
                    </div>
                  )}
                </div>
              </div>
            );
          })}
        </div>
      </div>

      {/* Légende */}
      <div style={{ marginTop: '1rem', padding: '12px 20px', display: 'flex', gap: '25px', fontSize: '0.8rem', background: 'var(--main-surface)', borderRadius: '10px', border: '1px solid var(--main-border)', boxShadow: '0 2px 5px rgba(0,0,0,0.05)', width: 'fit-content' }}>
        <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
          <div style={{ width: '30px', height: '12px', background: '#059669', borderRadius: '3px' }}></div>
          <span style={{ fontWeight: 600 }}>Phase</span>
        </div>
        <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
          <div style={{ width: '30px', height: '12px', background: '#3b82f6', borderRadius: '3px' }}></div>
          <span style={{ fontWeight: 600 }}>Tâche</span>
        </div>
        <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
          <div style={{ width: '12px', height: '12px', background: '#f59e0b', transform: 'rotate(45deg)', border: '1px solid #fff', boxShadow: '0 1px 3px rgba(0,0,0,0.2)' }}></div>
          <span style={{ fontWeight: 600, marginLeft: '4px' }}>Jalon</span>
        </div>
      </div>
    </div>
  );
};

const GanttRessourcesView = ({ resources, onChange, onAdd, onRemove }: { resources: any[], onChange: any, onAdd: any, onRemove: any }) => {
  return (
    <div style={{ flex: 1, background: 'var(--main-bg)', color: 'var(--main-text)', padding: '1.5rem', overflow: 'auto' }}>
      <div style={{ background: 'var(--main-surface)', padding: '2rem', borderRadius: '16px', border: '1px solid var(--main-border)', boxShadow: '0 4px 12px rgba(0,0,0,0.05)' }}>
        <div style={{ display: 'flex', alignItems: 'center', marginBottom: '1.5rem', gap: '0.8rem' }}>
          <h2 style={{ margin: 0, color: 'var(--table-primary)', fontWeight: 800 }}>👥 Ressources du planning</h2>
        </div>
        <div style={{ flex: 1, overflow: 'auto' }}>
          <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '0.85rem' }}>
            <thead style={{ background: 'var(--table-primary)', color: '#fff' }}>
              <tr>
                <th style={{ padding: '8px', border: '1px solid rgba(255,255,255,0.1)', width: '40px' }}>ID</th>
                <th style={{ padding: '8px', border: '1px solid rgba(255,255,255,0.1)', textAlign: 'left' }}>Nom</th>
                <th style={{ padding: '8px', border: '1px solid rgba(255,255,255,0.1)', textAlign: 'left' }}>Rôle</th>
                <th style={{ padding: '8px', border: '1px solid rgba(255,255,255,0.1)', textAlign: 'left' }}>Adresse email</th>
                <th style={{ padding: '8px', border: '1px solid rgba(255,255,255,0.1)', textAlign: 'left' }}>Téléphone</th>
                <th style={{ padding: '8px', border: '1px solid rgba(255,255,255,0.1)', width: '50px' }}>Action</th>
              </tr>
            </thead>
            <tbody>
              {resources.map((r: any, idx: number) => (
                <tr key={r.id} style={{ borderBottom: '1px solid var(--main-border)', background: idx % 2 === 0 ? 'transparent' : 'rgba(0,0,0,0.01)' }}>
                  <td style={{ padding: '8px', textAlign: 'center', fontWeight: 'bold' }}>{r.id}</td>
                  <td style={{ padding: '8px' }}>
                    <input value={r.name} onChange={(e) => onChange(r.id, 'name', e.target.value)} style={{ width: '100%', border: 'none', background: 'transparent', color: 'var(--main-text)', fontWeight: 600 }} />
                  </td>
                  <td style={{ padding: '8px' }}>
                    <input value={r.role} onChange={(e) => onChange(r.id, 'role', e.target.value)} style={{ width: '100%', border: 'none', background: 'transparent', color: 'var(--main-text)' }} />
                  </td>
                  <td style={{ padding: '8px' }}>
                    <input value={r.email} onChange={(e) => onChange(r.id, 'email', e.target.value)} style={{ width: '100%', border: 'none', background: 'transparent', color: '#3b82f6' }} />
                  </td>
                  <td style={{ padding: '8px' }}>
                    <input value={r.tel || ''} onChange={(e) => onChange(r.id, 'tel', e.target.value)} style={{ width: '100%', border: 'none', background: 'transparent', color: 'var(--main-text)' }} />
                  </td>
                  <td style={{ padding: '8px', textAlign: 'center' }}>
                    <button onClick={() => onRemove(r.id)} style={{ background: 'none', border: 'none', cursor: 'pointer' }}>🗑️</button>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
        <button onClick={onAdd} style={{ marginTop: '1rem', padding: '10px 20px', background: 'var(--table-primary)', color: '#fff', border: 'none', borderRadius: '8px', cursor: 'pointer', fontWeight: 800 }}>
          + Ajouter une ressource
        </button>
      </div>
    </div>
  );
};

const GanttParametrageView = () => {
  return (
    <div style={{ flex: 1, background: 'var(--main-bg)', color: 'var(--main-text)', padding: '2rem' }}>
      <div style={{ background: 'var(--main-surface)', padding: '2.5rem', borderRadius: '24px', border: '1px solid var(--main-border)', maxWidth: '800px', margin: '0 auto', boxShadow: '0 12px 40px rgba(0,0,0,0.1)' }}>
        <div style={{ display: 'flex', alignItems: 'center', marginBottom: '2.5rem', gap: '1rem' }}>
          <div style={{ width: '40px', height: '40px', borderRadius: '12px', background: 'var(--table-primary)', display: 'flex', alignItems: 'center', justifyContent: 'center', color: '#fff', fontSize: '1.2rem', fontWeight: 900 }}>⚙️</div>
          <h2 style={{ margin: 0, fontWeight: 900, fontSize: '1.8rem', color: 'var(--table-primary)' }}>Paramétrage Planning</h2>
        </div>
        <div style={{ display: 'flex', flexDirection: 'column', gap: '2rem' }}>
          <div style={{ background: 'rgba(0,0,0,0.02)', padding: '1.5rem', borderRadius: '12px', border: '1px solid var(--main-border)' }}>
            <label style={{ display: 'block', fontWeight: 800, marginBottom: '10px', fontSize: '0.9rem', color: 'var(--text-muted)' }}>MÉTHODE DE CALCUL</label>
            <select style={{ width: '100%', padding: '12px', borderRadius: '8px', border: '1px solid var(--main-border)', background: 'transparent', color: 'inherit', fontWeight: 600 }}>
              <option>Chemin critique (CPM)</option>
              <option>Agile / Scrum</option>
              <option>Manuel</option>
            </select>
          </div>
          <div style={{ background: 'rgba(0,0,0,0.02)', padding: '1.5rem', borderRadius: '12px', border: '1px solid var(--main-border)' }}>
            <label style={{ display: 'block', fontWeight: 800, marginBottom: '10px', fontSize: '0.9rem', color: 'var(--text-muted)' }}>UNITÉ DE CHARGE</label>
            <select style={{ width: '100%', padding: '12px', borderRadius: '8px', border: '1px solid var(--main-border)', background: 'transparent', color: 'inherit', fontWeight: 600 }}>
              <option>Jours-Homme (JH)</option>
              <option>Heures</option>
            </select>
          </div>
        </div>
      </div>
    </div>
  );
};


interface SelectionDossier {
  id: string;
  name: string;
  createdAt: number;
}

const SelectionDossiersView = ({
  projectId,
  fileId,
  onSelectDossier,
  onOpenDossier,
  currentDossierId,
  hideCreate,
  moduleType,
  title
}: {
  projectId: string;
  fileId?: string;
  onSelectDossier: (id: string, autoNavigate?: boolean) => void;
  onOpenDossier?: (id: string) => void;
  currentDossierId: string | null;
  hideCreate?: boolean;
  moduleType?: 'selection' | 'faisabilite' | string;
  title?: string;
}) => {
  const actualFileId = fileId || (moduleType === 'faisabilite' ? 'matrice-faisabilite' : moduleType === 'selection' ? 'selection-projets' : 'global');
  const storageKey = `faab-dossiers-${projectId}-${actualFileId}`;

  const getDefaultDossiers = (): SelectionDossier[] => {
    const legacyKey = actualFileId === 'matrice-faisabilite' ? `faisabilite-dossiers-${projectId}` : actualFileId === 'selection-projets' ? `selection-dossiers-${projectId}` : null;
    if (legacyKey) {
      const savedLegacy = localStorage.getItem(legacyKey);
      if (savedLegacy) {
        try {
          const parsed = JSON.parse(savedLegacy);
          if (Array.isArray(parsed) && parsed.length > 0) return parsed;
        } catch (e) {}
      }
    }
    if (actualFileId === 'matrice-faisabilite') {
      return [
        { id: 'doss-fais-1', name: "📁 Faisabilité Technique 2026", createdAt: Date.now() - 86400000 },
        { id: 'doss-fais-2', name: "📂 Étude de Marché Produit X", createdAt: Date.now() - 172800000 }
      ];
    }
    if (actualFileId === 'suivi-actions') {
      return [
        { id: 'doss-act-1', name: "📁 Actions Opérationnelles Q1", createdAt: Date.now() - 86400000 },
        { id: 'doss-act-2', name: "📂 Actions Correctives COPIL", createdAt: Date.now() - 172800000 }
      ];
    }
    if (actualFileId === 'suivi-livrables') {
      return [
        { id: 'doss-livr-1', name: "📁 Livrables Phase Exécution", createdAt: Date.now() - 86400000 },
        { id: 'doss-livr-2', name: "📂 Livrables Clôture & RUN", createdAt: Date.now() - 172800000 }
      ];
    }
    return [
      { id: `doss-${actualFileId}-1`, name: "📁 Dossier Principal (Session A)", createdAt: Date.now() - 86400000 },
      { id: `doss-${actualFileId}-2`, name: "📂 Historique & Variantes", createdAt: Date.now() - 172800000 }
    ];
  };

  const [dossiers, setDossiers] = useState<SelectionDossier[]>(() => {
    const saved = localStorage.getItem(storageKey);
    if (saved !== null) {
      try {
        const parsed = JSON.parse(saved);
        if (Array.isArray(parsed)) return parsed;
      } catch (e) {}
    }
    const def = getDefaultDossiers();
    localStorage.setItem(storageKey, JSON.stringify(def));
    return def;
  });

  const [newName, setNewName] = useState('');
  const [editingId, setEditingId] = useState<string | null>(null);
  const [editText, setEditText] = useState('');
  const [toast, setToast] = useState<string | null>(null);
  const [dossierToDelete, setDossierToDelete] = useState<{ id: string; name: string } | null>(null);

  // Sync state whenever storageKey changes
  useEffect(() => {
    const saved = localStorage.getItem(storageKey);
    if (saved !== null) {
      try {
        const parsed = JSON.parse(saved);
        if (Array.isArray(parsed)) {
          setDossiers(parsed);
          return;
        }
      } catch (e) {}
    }
    const def = getDefaultDossiers();
    localStorage.setItem(storageKey, JSON.stringify(def));
    setDossiers(def);
  }, [storageKey]);

  const saveDossiers = (list: SelectionDossier[]) => {
    const sorted = [...list].sort((a, b) => b.createdAt - a.createdAt);
    setDossiers(sorted);
    localStorage.setItem(storageKey, JSON.stringify(sorted));
    if (actualFileId === 'selection-projets') {
      localStorage.setItem(`selection-dossiers-${projectId}`, JSON.stringify(sorted));
    } else if (actualFileId === 'matrice-faisabilite') {
      localStorage.setItem(`faisabilite-dossiers-${projectId}`, JSON.stringify(sorted));
    }
  };

  const showToast = (msg: string) => {
    setToast(msg);
    setTimeout(() => setToast(null), 3500);
  };

  const createDossier = () => {
    let name = newName.trim();
    if (!name) {
      name = `Nouveau Dossier ${dossiers.length + 1}`;
    }
    if (!name.startsWith('📁') && !name.startsWith('📂')) {
      name = `📁 ${name}`;
    }
    const newDossier: SelectionDossier = {
      id: `dossier-${actualFileId}-${Date.now()}`,
      name,
      createdAt: Date.now()
    };
    const updated = [newDossier, ...dossiers];
    saveDossiers(updated);
    setNewName('');

    // Select the new dossier without redirecting away so user can see it
    onSelectDossier(newDossier.id, false);
    showToast(`✅ Dossier "${name}" créé et activé avec succès !`);
  };

  const duplicateDossier = (d: SelectionDossier) => {
    const newDossier: SelectionDossier = {
      id: `dossier-${actualFileId}-${Date.now()}`,
      name: `${d.name} (Copie)`,
      createdAt: Date.now()
    };
    try {
      for (let i = 0; i < localStorage.length; i++) {
        const k = localStorage.key(i);
        if (k && k.includes(`-${d.id}`)) {
          const newKey = k.replace(`-${d.id}`, `-${newDossier.id}`);
          localStorage.setItem(newKey, localStorage.getItem(k) || '');
        }
      }
    } catch (e) {
      console.error(e);
    }

    const updated = [newDossier, ...dossiers];
    saveDossiers(updated);
    showToast(`📋 Dossier "${d.name}" dupliqué avec succès !`);
  };

  const renameDossier = (id: string) => {
    if (!editText.trim()) return;
    const updated = dossiers.map(d => d.id === id ? { ...d, name: editText.trim() } : d);
    saveDossiers(updated);
    setEditingId(null);
    showToast(`✏️ Dossier renommé en "${editText.trim()}".`);
  };

  const deleteDossier = (id: string, name: string) => {
    const updated = dossiers.filter(d => d.id !== id);
    saveDossiers(updated);
    if (currentDossierId === id) {
      if (updated.length > 0) {
        onSelectDossier(updated[0].id, false);
      } else {
        onSelectDossier('', false);
      }
    }
    try {
      const keysToRemove: string[] = [];
      for (let i = 0; i < localStorage.length; i++) {
        const k = localStorage.key(i);
        if (k && k.includes(`-${id}`)) {
          keysToRemove.push(k);
        }
      }
      keysToRemove.forEach(k => localStorage.removeItem(k));
    } catch (e) {
      console.error(e);
    }
    showToast(`🗑️ Dossier "${name}" supprimé.`);
  };

  const activeDossier = dossiers.find(d => d.id === currentDossierId) || (dossiers.length > 0 ? dossiers[0] : null);

  return (
    <div style={{ flex: 1, padding: '2rem', background: 'var(--main-bg)', color: 'var(--main-text)', overflow: 'auto' }}>
      <div style={{ maxWidth: '1080px', margin: '0 auto' }}>
        {/* Toast feedback */}
        {toast && (
          <div style={{
            position: 'fixed',
            top: '20px',
            right: '20px',
            background: 'linear-gradient(135deg, #10b981, #059669)',
            color: '#fff',
            padding: '12px 24px',
            borderRadius: '12px',
            fontWeight: 800,
            fontSize: '0.9rem',
            boxShadow: '0 10px 25px rgba(0,0,0,0.25)',
            zIndex: 9999,
            display: 'flex',
            alignItems: 'center',
            gap: '10px',
            animation: 'fadeIn 0.2s ease-out'
          }}>
            {toast}
          </div>
        )}

        {/* Header Summary */}
        <div style={{
          display: 'flex',
          justifyContent: 'space-between',
          alignItems: 'center',
          marginBottom: '1.5rem',
          background: 'var(--main-surface)',
          padding: '1.5rem 2.5rem',
          borderRadius: '24px',
          boxShadow: '0 8px 32px rgba(0,0,0,0.1)',
          border: '1px solid var(--main-border)',
          position: 'relative',
          overflow: 'hidden'
        }}>
          <div style={{ position: 'absolute', top: 0, right: 0, width: '200px', height: '100%', background: 'linear-gradient(90deg, transparent 0%, var(--table-primary) 150%)', opacity: 0.05 }} />
          <div>
            <h2 style={{ margin: 0, fontSize: '1.8rem', fontWeight: 900, color: 'var(--table-primary)', letterSpacing: '-0.5px' }}>
              {title || `Gestion des Dossiers : ${actualFileId}`}
            </h2>
            <p style={{ margin: '6px 0 0 0', fontSize: '0.9rem', color: 'var(--text-muted)', fontWeight: 500 }}>
              Créez, gérez vos sessions de travail, dupliquez et activez vos dossiers en toute réactivité
            </p>
          </div>
          <div style={{ display: 'flex', gap: '1rem', alignItems: 'center', background: 'rgba(0,0,0,0.03)', padding: '10px 20px', borderRadius: '16px', border: '1px solid var(--main-border)' }}>
            <span style={{ fontSize: '1.4rem' }}>📁</span>
            <span style={{ fontWeight: 900, fontSize: '1.2rem', color: 'var(--main-text)' }}>{dossiers.length} dossiers</span>
          </div>
        </div>

        {/* Active Dossier Highlight Banner */}
        {activeDossier && (
          <div style={{
            background: 'linear-gradient(135deg, rgba(16, 185, 129, 0.12), rgba(56, 189, 248, 0.08))',
            border: '1px solid rgba(16, 185, 129, 0.3)',
            borderRadius: '16px',
            padding: '1rem 1.8rem',
            marginBottom: '2rem',
            display: 'flex',
            justifyContent: 'space-between',
            alignItems: 'center',
            boxShadow: '0 4px 16px rgba(0,0,0,0.05)'
          }}>
            <div style={{ display: 'flex', alignItems: 'center', gap: '14px' }}>
              <div style={{ width: '42px', height: '42px', borderRadius: '12px', background: '#10b981', color: '#fff', display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: '1.3rem', boxShadow: '0 4px 12px rgba(16, 185, 129, 0.3)' }}>
                ⭐
              </div>
              <div>
                <span style={{ fontSize: '0.72rem', fontWeight: 800, color: '#10b981', textTransform: 'uppercase', letterSpacing: '0.5px' }}>Dossier actuellement actif</span>
                <div style={{ fontSize: '1.15rem', fontWeight: 900, color: 'var(--main-text)' }}>{activeDossier.name}</div>
              </div>
            </div>
            <button
              onClick={() => onOpenDossier ? onOpenDossier(activeDossier.id) : onSelectDossier(activeDossier.id, true)}
              style={{
                background: 'linear-gradient(135deg, #10b981, #059669)',
                color: '#fff',
                border: 'none',
                padding: '10px 22px',
                borderRadius: '12px',
                fontWeight: 800,
                fontSize: '0.85rem',
                cursor: 'pointer',
                display: 'flex',
                alignItems: 'center',
                gap: '8px',
                boxShadow: '0 4px 14px rgba(16, 185, 129, 0.3)',
                transition: 'transform 0.15s ease'
              }}
              onMouseEnter={(e) => e.currentTarget.style.transform = 'translateY(-2px)'}
              onMouseLeave={(e) => e.currentTarget.style.transform = 'translateY(0)'}
            >
              <span>🚀</span> Accéder aux données du classeur →
            </button>
          </div>
        )}

        {/* Create Folder Form */}
        {!hideCreate && (
          <div style={{
            background: 'var(--main-surface)',
            padding: '1.8rem 2rem',
            borderRadius: '24px',
            marginBottom: '2.5rem',
            border: '1px solid var(--main-border)',
            boxShadow: '0 4px 20px rgba(0,0,0,0.03)',
            display: 'flex',
            gap: '1.25rem',
            alignItems: 'center'
          }}>
            <div style={{ flex: 1, position: 'relative' }}>
              <input
                placeholder="Nom du nouveau dossier (ex: Projet Alpha 2026, Session Q2, Variante A)..."
                value={newName}
                onChange={(e) => setNewName(e.target.value)}
                onKeyDown={(e) => e.key === 'Enter' && createDossier()}
                style={{
                  width: '100%',
                  padding: '16px 24px',
                  borderRadius: '16px',
                  border: '2px solid var(--main-border)',
                  background: 'rgba(0,0,0,0.02)',
                  color: 'var(--main-text)',
                  fontSize: '1rem',
                  fontWeight: 600,
                  outline: 'none',
                  transition: 'all 0.3s cubic-bezier(0.4, 0, 0.2, 1)'
                }}
                onFocus={(e) => e.currentTarget.style.borderColor = 'var(--table-primary)'}
                onBlur={(e) => e.currentTarget.style.borderColor = 'var(--main-border)'}
              />
            </div>
            <button
              onClick={createDossier}
              style={{
                background: 'linear-gradient(135deg, var(--table-primary) 0%, #469b6f 100%)',
                color: '#fff',
                border: 'none',
                padding: '16px 36px',
                borderRadius: '16px',
                fontWeight: 800,
                cursor: 'pointer',
                boxShadow: '0 8px 16px rgba(70, 155, 111, 0.25)',
                display: 'flex',
                alignItems: 'center',
                gap: '10px',
                fontSize: '1rem',
                transition: 'all 0.3s ease',
                whiteSpace: 'nowrap'
              }}
              onMouseEnter={(e) => { e.currentTarget.style.transform = 'translateY(-3px)'; e.currentTarget.style.boxShadow = '0 12px 24px rgba(70, 155, 111, 0.35)'; }}
              onMouseLeave={(e) => { e.currentTarget.style.transform = 'translateY(0)'; e.currentTarget.style.boxShadow = '0 8px 16px rgba(70, 155, 111, 0.25)'; }}
            >
              <span style={{ fontSize: '1.2rem', lineHeight: '1' }}>+</span> Créer le dossier
            </button>
          </div>
        )}

        {/* Dossiers Grid */}
        <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fill, minmax(320px, 1fr))', gap: '2rem' }}>
          {dossiers.map(d => {
            const isActive = (currentDossierId === d.id) || (!currentDossierId && activeDossier?.id === d.id);
            return (
              <div
                key={d.id}
                onClick={() => {
                  if (editingId !== d.id) {
                    onSelectDossier(d.id, false);
                    showToast(`⭐ Dossier "${d.name}" défini comme actif !`);
                  }
                }}
                style={{
                  background: 'var(--main-surface)',
                  padding: '1.75rem',
                  borderRadius: '24px',
                  cursor: 'pointer',
                  position: 'relative',
                  transition: 'all 0.3s cubic-bezier(0.175, 0.885, 0.32, 1.275)',
                  border: `3px solid ${isActive ? 'var(--table-primary)' : 'var(--main-border)'}`,
                  boxShadow: isActive ? '0 12px 30px rgba(70, 155, 111, 0.2)' : '0 4px 15px rgba(0,0,0,0.05)',
                  display: 'flex',
                  flexDirection: 'column',
                  gap: '1.25rem',
                  overflow: 'hidden'
                }}
                onMouseEnter={(e) => {
                  if (!isActive) {
                    e.currentTarget.style.boxShadow = '0 8px 25px rgba(0,0,0,0.1)';
                    e.currentTarget.style.transform = 'translateY(-6px)';
                  }
                }}
                onMouseLeave={(e) => {
                  if (!isActive) {
                    e.currentTarget.style.boxShadow = '0 4px 15px rgba(0,0,0,0.05)';
                    e.currentTarget.style.transform = 'translateY(0)';
                  }
                }}
              >
                <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start' }}>
                  <div style={{
                    width: '56px',
                    height: '56px',
                    borderRadius: '18px',
                    background: isActive ? 'var(--table-primary)' : 'rgba(0,0,0,0.04)',
                    color: isActive ? '#fff' : 'inherit',
                    display: 'flex',
                    alignItems: 'center',
                    justifyContent: 'center',
                    fontSize: '1.6rem',
                    boxShadow: isActive ? 'inset 0 2px 4px rgba(255,255,255,0.2)' : 'none'
                  }}>
                    {isActive ? '📂' : '📁'}
                  </div>
                  {isActive ? (
                    <div style={{
                      background: 'var(--table-primary)',
                      color: '#fff',
                      padding: '5px 12px',
                      borderRadius: '50px',
                      fontSize: '0.7rem',
                      fontWeight: 900,
                      letterSpacing: '1px',
                      boxShadow: '0 4px 10px rgba(70, 155, 111, 0.3)',
                    }}>
                      ACTIF
                    </div>
                  ) : (
                    <button
                      onClick={(e) => {
                        e.stopPropagation();
                        onSelectDossier(d.id, false);
                        showToast(`⭐ Dossier "${d.name}" défini comme actif !`);
                      }}
                      style={{
                        background: 'transparent',
                        border: '1px solid var(--main-border)',
                        color: 'var(--text-muted)',
                        padding: '4px 10px',
                        borderRadius: '20px',
                        fontSize: '0.7rem',
                        fontWeight: 700,
                        cursor: 'pointer'
                      }}
                    >
                      Activer
                    </button>
                  )}
                </div>

                <div style={{ flex: 1 }}>
                  {editingId === d.id ? (
                    <div style={{ display: 'flex', gap: '0.5rem', width: '100%' }}>
                      <input
                        autoFocus
                        value={editText}
                        onChange={(e) => setEditText(e.target.value)}
                        onClick={(e) => e.stopPropagation()}
                        onKeyDown={(e) => e.key === 'Enter' && renameDossier(d.id)}
                        style={{
                          flex: 1,
                          padding: '10px',
                          borderRadius: '10px',
                          border: '2px solid var(--table-primary)',
                          background: 'var(--main-bg)',
                          color: 'var(--main-text)',
                          fontSize: '1rem',
                          fontWeight: 700
                        }}
                      />
                      <button
                        onClick={(e) => { e.stopPropagation(); renameDossier(d.id); }}
                        style={{ background: 'var(--table-primary)', color: '#fff', border: 'none', borderRadius: '10px', padding: '0 15px', fontWeight: 800, cursor: 'pointer' }}
                      >
                        OK
                      </button>
                    </div>
                  ) : (
                    <div>
                      <div style={{ fontWeight: 900, fontSize: '1.15rem', color: 'var(--main-text)', marginBottom: '6px', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>
                        {d.name}
                      </div>
                      <div style={{ fontSize: '0.8rem', color: 'var(--text-muted)', display: 'flex', alignItems: 'center', gap: '6px' }}>
                        <span style={{ opacity: 0.6 }}>📅</span> {new Date(d.createdAt).toLocaleDateString(undefined, { day: 'numeric', month: 'short', year: 'numeric' })}
                      </div>
                    </div>
                  )}
                </div>

                {/* Primary Action Button */}
                <button
                  onClick={(e) => {
                    e.stopPropagation();
                    if (onOpenDossier) onOpenDossier(d.id);
                    else onSelectDossier(d.id, true);
                  }}
                  style={{
                    width: '100%',
                    padding: '10px 14px',
                    borderRadius: '12px',
                    background: isActive ? 'linear-gradient(135deg, #10b981, #059669)' : 'rgba(255, 255, 255, 0.06)',
                    color: isActive ? '#fff' : 'var(--main-text)',
                    border: isActive ? 'none' : '1px solid var(--main-border)',
                    fontWeight: 800,
                    fontSize: '0.85rem',
                    cursor: 'pointer',
                    display: 'flex',
                    alignItems: 'center',
                    justifyContent: 'center',
                    gap: '8px',
                    boxShadow: isActive ? '0 4px 12px rgba(16, 185, 129, 0.25)' : 'none',
                    transition: 'all 0.2s ease'
                  }}
                  onMouseEnter={(e) => {
                    e.currentTarget.style.transform = 'translateY(-2px)';
                    if (!isActive) e.currentTarget.style.background = 'rgba(255, 255, 255, 0.12)';
                  }}
                  onMouseLeave={(e) => {
                    e.currentTarget.style.transform = 'translateY(0)';
                    if (!isActive) e.currentTarget.style.background = 'rgba(255, 255, 255, 0.06)';
                  }}
                >
                  <span>🚀</span> Ouvrir ce dossier
                </button>

                <div style={{ display: 'flex', gap: '0.5rem', borderTop: '1px solid var(--main-border)', paddingTop: '1rem', marginTop: '0.2rem' }}>
                  <button
                    onClick={(e) => { e.stopPropagation(); duplicateDossier(d); }}
                    style={{
                      flex: 1,
                      background: 'rgba(0,0,0,0.03)',
                      border: 'none',
                      borderRadius: '10px',
                      padding: '8px',
                      color: 'var(--text-muted)',
                      fontSize: '0.78rem',
                      fontWeight: 700,
                      cursor: 'pointer',
                      transition: 'all 0.2s',
                      display: 'flex',
                      alignItems: 'center',
                      justifyContent: 'center',
                      gap: '4px'
                    }}
                    onMouseEnter={(e) => { e.currentTarget.style.background = 'rgba(0,0,0,0.08)'; e.currentTarget.style.color = 'var(--main-text)'; }}
                    onMouseLeave={(e) => { e.currentTarget.style.background = 'rgba(0,0,0,0.03)'; e.currentTarget.style.color = 'var(--text-muted)'; }}
                  >
                    <span>📋</span> Copier
                  </button>
                  <button
                    onClick={(e) => { e.stopPropagation(); setEditingId(d.id); setEditText(d.name); }}
                    style={{
                      flex: 1,
                      background: 'rgba(0,0,0,0.03)',
                      border: 'none',
                      borderRadius: '10px',
                      padding: '8px',
                      color: 'var(--text-muted)',
                      fontSize: '0.78rem',
                      fontWeight: 700,
                      cursor: 'pointer',
                      transition: 'all 0.2s',
                      display: 'flex',
                      alignItems: 'center',
                      justifyContent: 'center',
                      gap: '4px'
                    }}
                    onMouseEnter={(e) => { e.currentTarget.style.background = 'rgba(0,0,0,0.08)'; e.currentTarget.style.color = 'var(--main-text)'; }}
                    onMouseLeave={(e) => { e.currentTarget.style.background = 'rgba(0,0,0,0.03)'; e.currentTarget.style.color = 'var(--text-muted)'; }}
                  >
                    <span>✏️</span> Renommer
                  </button>
                  <button
                    onClick={(e) => {
                      e.stopPropagation();
                      setDossierToDelete({ id: d.id, name: d.name });
                    }}
                    style={{
                      width: '42px',
                      background: 'rgba(239, 68, 68, 0.08)',
                      border: 'none',
                      borderRadius: '10px',
                      padding: '8px',
                      color: '#ef4444',
                      fontSize: '1.1rem',
                      cursor: 'pointer',
                      transition: 'all 0.2s',
                      display: 'flex',
                      alignItems: 'center',
                      justifyContent: 'center'
                    }}
                    onMouseEnter={(e) => { e.currentTarget.style.background = 'rgba(239, 68, 68, 0.2)'; e.currentTarget.style.transform = 'scale(1.1)'; }}
                    onMouseLeave={(e) => { e.currentTarget.style.background = 'rgba(239, 68, 68, 0.08)'; e.currentTarget.style.transform = 'scale(1)'; }}
                    title="Supprimer ce dossier"
                  >
                    🗑️
                  </button>
                </div>
              </div>
            );
          })}

          {dossiers.length === 0 && (
            <div style={{
              gridColumn: '1 / -1',
              textAlign: 'center',
              padding: '5rem 2rem',
              background: 'var(--main-surface)',
              borderRadius: '32px',
              border: '2px dashed var(--main-border)',
              color: 'var(--text-muted)',
              display: 'flex',
              flexDirection: 'column',
              alignItems: 'center',
              gap: '1rem'
            }}>
              <div style={{ fontSize: '4rem', filter: 'grayscale(1)', opacity: 0.5 }}>📂</div>
              <div style={{ fontSize: '1.4rem', fontWeight: 900, color: 'var(--main-text)' }}>Aucun dossier disponible</div>
              <div style={{ fontSize: '1rem', maxWidth: '400px', lineHeight: '1.6' }}>
                Créez votre premier dossier de travail ci-dessus pour commencer à organiser vos analyses et données.
              </div>
            </div>
          )}
        </div>

        {/* Modal de Confirmation de Suppression de Dossier */}
        {dossierToDelete && (
          <div style={{
            position: 'fixed',
            top: 0,
            left: 0,
            right: 0,
            bottom: 0,
            background: 'rgba(0, 0, 0, 0.7)',
            backdropFilter: 'blur(6px)',
            display: 'flex',
            alignItems: 'center',
            justifyContent: 'center',
            zIndex: 10000,
            animation: 'fadeIn 0.2s ease-out'
          }}>
            <div style={{
              background: 'var(--main-surface)',
              border: '2px solid rgba(239, 68, 68, 0.4)',
              borderRadius: '24px',
              padding: '2.5rem',
              maxWidth: '480px',
              width: '90%',
              boxShadow: '0 25px 60px rgba(0,0,0,0.5)',
              textAlign: 'center'
            }}>
              <div style={{ fontSize: '3.5rem', marginBottom: '1rem' }}>⚠️</div>
              <h3 style={{ margin: '0 0 0.75rem 0', fontSize: '1.45rem', fontWeight: 900, color: 'var(--main-text)' }}>
                Supprimer le dossier ?
              </h3>
              <p style={{ margin: '0 0 2rem 0', color: 'var(--text-muted)', fontSize: '0.95rem', lineHeight: '1.6' }}>
                Êtes-vous certain de vouloir supprimer le dossier <strong style={{ color: '#ef4444' }}>« {dossierToDelete.name} »</strong> ?<br />
                Toutes les données associées à cette session seront définitivement effacées.
              </p>
              <div style={{ display: 'flex', gap: '1rem', justifyContent: 'center' }}>
                <button
                  onClick={() => setDossierToDelete(null)}
                  style={{
                    padding: '12px 24px',
                    borderRadius: '12px',
                    border: '1px solid var(--main-border)',
                    background: 'rgba(255, 255, 255, 0.05)',
                    color: 'var(--main-text)',
                    fontWeight: 700,
                    cursor: 'pointer',
                    fontSize: '0.95rem'
                  }}
                >
                  Annuler
                </button>
                <button
                  onClick={() => {
                    deleteDossier(dossierToDelete.id, dossierToDelete.name);
                    setDossierToDelete(null);
                  }}
                  style={{
                    padding: '12px 28px',
                    borderRadius: '12px',
                    border: 'none',
                    background: 'linear-gradient(135deg, #ef4444, #b91c1c)',
                    color: '#fff',
                    fontWeight: 800,
                    cursor: 'pointer',
                    fontSize: '0.95rem',
                    boxShadow: '0 6px 20px rgba(239, 68, 68, 0.4)',
                    display: 'flex',
                    alignItems: 'center',
                    gap: '8px'
                  }}
                >
                  <span>🗑️</span> Confirmer la suppression
                </button>
              </div>
            </div>
          </div>
        )}
      </div>
    </div>
  );
};

const SelectionEvaluationView = ({ fileId, projectId, currentDossierId, initialView, onDossierChange }: {
  fileId: string;
  projectId: string;
  currentDossierId: string | null;
  initialView?: 'evaluation' | 'criteria' | 'dashboard' | 'database' | 'dossiers' | 'settings',
  onDossierChange?: (id: string) => void
}) => {
  const [viewMode, setViewMode] = useState<'evaluation' | 'criteria' | 'dashboard' | 'database' | 'dossiers' | 'settings'>(initialView || 'evaluation');

  useEffect(() => {
    if (initialView) setViewMode(initialView);
  }, [initialView]);
  const dossierSubKey = currentDossierId ? `-${currentDossierId}` : "";
  const storageKey = `selection-projets-${projectId}-${fileId}${dossierSubKey}`;

  const [criteria, setCriteria] = useState(() => {
    const saved = localStorage.getItem(`selection-criteria-${projectId}`);
    return saved ? JSON.parse(saved) : [
      { code: 'C1', name: 'Degré de cohérence avec la mission de l\'entreprise', weight: 4, options: ["Fortement incohérente", "Un peu incohérente", "Neutre", "Cohérente", "Renforce la mission"] },
      { code: 'C2', name: 'Degré de cohérence et d\'alignement avec la stratégie', weight: 3, options: ["Fortement incohérente", "Un peu incohérente", "Neutre", "Cohérente", "Appui la stratégie (deux axes ou plus)"] },
      { code: 'C3', name: 'Degré d\'urgence du projet', weight: 3, options: ["Sans délais précis", "Flexible", "Normal", "Assez pressé", "Très urgent"] },
      { code: 'C4', name: 'Stabilité des objectifs du projet', weight: 3, options: ["Chaotiques", "Inconstants", "Risques modérés", "Relativement stable", "Très stable"] },
      { code: 'C5', name: 'Soutien des parties prenantes', weight: 2, options: ["Fortes oppositions", "Quelques oppositions", "Equilibre", "Bon soutien", "Soutien inconditionnel"] },
      { code: 'C6', name: 'Durée du projet', weight: 3, options: ["Très long > 3 ans", "Long de 18 mois à 3 ans", "Normal de 12 à 18 mois", "Court de 3 à 12 mois", "Très court < 3 mois"] },
      { code: 'C7', name: 'Coût du projet (Capex + Opex externes)', weight: 2, options: ["Très élevé > 1,2 G FCFA", "Elevé de 600 M à 1,2 G FCFA", "Moyen de 120 M à 600 M FCFA", "Bon marché < 1,2 M FCFA", "Presque sans coûts"] },
      { code: 'C8', name: 'Rapport coût x bénéfice', weight: 4, options: ["Très coûteux et peu de résultats", "Coûteux et peu de résultats", "Equilibre raisonnable", "Relation favorable", "Relation très favorable"] },
      { code: 'C9', name: 'Disponibilité des ressources', weight: 1, options: ["Ressources très concurrencées", "Exige quelques remaniements", "Peuvent être allouées", "Facilement disponibles", "Disponibles et sans attributions"] },
      { code: 'C10', name: 'Degré de risque du projet durant son exécution', weight: 2, options: ["Très risqué", "Risque moyen", "Peu de risque", "Presque sans risque", "Sans aucun risque"] },
      { code: 'C11', name: 'Degré de complexité du projet', weight: 3, options: ["Très complexe", "Complexe", "Moyennement complexe", "Assez simple", "Très simple"] },
    ];
  });

  const [projectNames, setProjectNames] = useState<string[]>(() => {
    const saved = localStorage.getItem(`${storageKey}-names`);
    if (saved) return JSON.parse(saved);
    return ["Projet Rénovation", "Digitalisation Flux", "Audit Performance", "Expansion Régionale", "Nouveau CRM", "Optimisation Stocks", "Formation Cadres", "Marketing Viral"];
  });

  const [data, setData] = useState<string[][]>(() => {
    const saved = localStorage.getItem(storageKey);
    if (saved) return JSON.parse(saved);
    // Initialize with some random values for "contents" (exemples avec contenus)
    return Array(8).fill(null).map(() =>
      criteria.map((c: any) => c.options[Math.floor(Math.random() * c.options.length)])
    );
  });

  useEffect(() => {
    localStorage.setItem(storageKey, JSON.stringify(data));
  }, [data, storageKey]);

  useEffect(() => {
    localStorage.setItem(`${storageKey}-names`, JSON.stringify(projectNames));
  }, [projectNames, storageKey]);

  // Sync data when storageKey changes (switching dossiers)
  useEffect(() => {
    const savedNames = localStorage.getItem(`${storageKey}-names`);
    if (savedNames) {
      setProjectNames(JSON.parse(savedNames));
    } else {
      setProjectNames(["Projet Rénovation", "Digitalisation Flux", "Audit Performance", "Expansion Régionale", "Nouveau CRM", "Optimisation Stocks", "Formation Cadres", "Marketing Viral"]);
    }
  }, [storageKey]);

  useEffect(() => {
    const savedData = localStorage.getItem(storageKey);
    if (savedData) {
      setData(JSON.parse(savedData));
    } else {
      setData(Array(8).fill(null).map(() =>
        criteria.map((c: any) => c.options[Math.floor(Math.random() * c.options.length)])
      ));
    }
  }, [storageKey]);

  useEffect(() => {
    localStorage.setItem(`selection-criteria-${projectId}`, JSON.stringify(criteria));
  }, [criteria, projectId]);

  const handleProjectNameChange = (idx: number, name: string) => {
    const newNames = [...projectNames];
    newNames[idx] = name;
    setProjectNames(newNames);
  };

  const handleChange = (pIdx: number, cIdx: number, val: string) => {
    const newData = [...data];
    newData[pIdx] = [...newData[pIdx]];
    newData[pIdx][cIdx] = val;
    setData(newData);
  };

  const handleCriteriaChange = (idx: number, field: string, val: any) => {
    const newCriteria = [...criteria];
    newCriteria[idx] = { ...newCriteria[idx], [field]: val };
    setCriteria(newCriteria);
  };

  const scores = projectNames.map((_, pIdx) => {
    let score = 0;
    criteria.forEach((c: any, cIdx: number) => {
      const val = data[pIdx][cIdx];
      if (val) {
        const optionIdx = c.options.indexOf(val);
        if (optionIdx !== -1) {
          const rating = optionIdx - 2; // -2, -1, 0, 1, 2
          score += rating * c.weight;
        }
      }
    });
    return score;
  });

  const getPriority = (score: number) => {
    if (score >= 32) return { p: 'P1', c: '#e8f5e9', t: '#2e7d32', label: 'Stricte priorité (P1)' };
    if (score >= 0) return { p: 'P2', c: '#fff3e0', t: '#e65100', label: 'Priorité élevée (P2)' };
    if (score >= -32) return { p: 'P3', c: '#fbe9e7', t: '#d84315', label: 'Priorité normale (P3)' };
    return { p: 'P4', c: '#ffebee', t: '#c62828', label: 'Priorité faible (P4)' };
  };

  return (
    <div style={{ flex: 1, display: 'flex', flexDirection: 'column', backgroundColor: 'var(--main-bg)', color: 'var(--main-text)', overflow: 'hidden' }}>
      <div style={{ flex: 1, overflow: 'auto', padding: '1.5rem' }}>
        {viewMode === 'database' && (
          <div style={{ background: 'var(--main-surface)', borderRadius: '12px', padding: '1.5rem', boxShadow: '0 4px 6px -1px rgba(0,0,0,0.1)', border: '1px solid var(--main-border)' }}>
            <table style={{ width: '100%', borderCollapse: 'collapse', borderRadius: '8px', overflow: 'hidden' }}>
              <thead style={{ background: 'var(--table-primary)', color: '#fff' }}>
                <tr>
                  <th style={{ padding: '12px', textAlign: 'left' }}>#</th>
                  <th style={{ padding: '12px', textAlign: 'left' }}>Nom du Projet</th>
                  <th style={{ padding: '12px', textAlign: 'left' }}>Date de création</th>
                  <th style={{ padding: '12px', textAlign: 'center' }}>Actions</th>
                </tr>
              </thead>
              <tbody>
                {projectNames.map((p, i) => (
                  <tr key={i} style={{ borderBottom: '1px solid var(--main-border)', background: i % 2 === 0 ? 'rgba(0,0,0,0.02)' : 'transparent' }}>
                    <td style={{ padding: '12px', fontWeight: 700 }}>{i + 1}</td>
                    <td style={{ padding: '12px' }}>
                      <input
                        value={p}
                        onChange={(e) => handleProjectNameChange(i, e.target.value)}
                        style={{ width: '100%', padding: '8px', borderRadius: '4px', border: '1px solid var(--main-border)', background: 'rgba(0,0,0,0.02)', color: 'var(--main-text)', fontWeight: 600 }}
                      />
                    </td>
                    <td style={{ padding: '12px', color: 'var(--text-muted)', fontSize: '0.8rem' }}>07/03/2026</td>
                    <td style={{ padding: '12px', textAlign: 'center' }}>
                      <button
                        onClick={() => {
                          const newNames = projectNames.filter((_, idx) => idx !== i);
                          const newData = data.filter((_, idx) => idx !== i);
                          setProjectNames(newNames);
                          setData(newData);
                        }}
                        style={{ background: 'none', border: 'none', cursor: 'pointer', fontSize: '1.2rem', transition: 'transform 0.15s' }}
                        onMouseEnter={(e) => e.currentTarget.style.transform = 'scale(1.2)'}
                        onMouseLeave={(e) => e.currentTarget.style.transform = 'scale(1)'}
                        title="Supprimer ce projet"
                      >
                        🗑️
                      </button>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
            <button
              onClick={() => {
                setProjectNames([...projectNames, `Nouveau Projet ${projectNames.length + 1}`]);
                setData([...data, Array(criteria.length).fill('')]);
              }}
              style={{ marginTop: '1.5rem', padding: '10px 24px', background: 'var(--table-primary)', color: '#fff', border: 'none', borderRadius: '8px', cursor: 'pointer', fontWeight: 800, fontSize: '0.9rem', boxShadow: '0 4px 12px rgba(59, 130, 246, 0.3)' }}
            >
              + Ajouter un projet
            </button>

          </div>
        )}

        {viewMode === 'settings' && (
          <div style={{ background: 'var(--main-surface)', borderRadius: '12px', padding: '1.5rem', boxShadow: '0 4px 6px -1px rgba(0,0,0,0.1)', border: '1px solid var(--main-border)' }}>
            <div style={{ padding: '1rem', background: 'rgba(0,0,0,0.02)', borderRadius: '8px' }}>
              <p>Configurez ici les options globales pour la sélection des projets.</p>
              <div style={{ marginTop: '1rem', display: 'flex', gap: '2rem' }}>
                <label style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
                  <input type="checkbox" defaultChecked /> Activer les alertes de priorité
                </label>
                <label style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
                  <input type="checkbox" defaultChecked /> Afficher les coefficients dans l'export
                </label>
              </div>
            </div>
          </div>
        )}

        {viewMode === 'evaluation' && (
          <div style={{ background: 'var(--main-surface)', borderRadius: '12px', padding: '1.5rem', boxShadow: '0 4px 6px -1px rgba(0,0,0,0.1)', border: '1px solid var(--main-border)' }}>
            <div style={{ overflowX: 'auto' }}>
              <table style={{ minWidth: '1000px', width: '100%', borderCollapse: 'collapse', fontSize: '0.85rem' }}>
                <thead>
                  <tr style={{ backgroundColor: 'var(--table-primary)', color: '#fff' }}>
                    <th style={{ padding: '12px', width: '60px', border: '1px solid rgba(255,255,255,0.1)' }}>Code</th>
                    <th style={{ padding: '12px', textAlign: 'left', border: '1px solid rgba(255,255,255,0.1)' }}>Critères de sélection</th>
                    <th style={{ padding: '12px', width: '80px', border: '1px solid rgba(255,255,255,0.1)' }}>Coeff.</th>
                    {projectNames.map((p, i) => (
                      <th key={i} style={{ padding: '12px', minWidth: '220px', border: '1px solid rgba(255,255,255,0.1)' }}>
                        <textarea
                          value={p}
                          onChange={(e) => handleProjectNameChange(i, e.target.value)}
                          rows={3}
                          style={{
                            width: '100%',
                            border: 'none',
                            background: 'rgba(255,255,255,0.15)',
                            color: '#fff',
                            textAlign: 'center',
                            fontSize: '0.85rem',
                            fontWeight: 800,
                            padding: '8px',
                            borderRadius: '6px',
                            outline: 'none',
                            resize: 'none',
                            fontFamily: 'inherit',
                            lineHeight: '1.2'
                          }}
                        />
                      </th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {criteria.map((c: any, cIdx: number) => (
                    <tr key={cIdx} style={{ borderBottom: '1px solid var(--main-border)', background: cIdx % 2 === 0 ? 'rgba(0,0,0,0.02)' : 'transparent' }}>
                      <td style={{ padding: '12px', textAlign: 'center', fontWeight: 'bold', color: 'var(--main-text)' }}>{c.code}</td>
                      <td style={{
                        padding: '12px',
                        color: 'var(--main-text)',
                        fontWeight: 500,
                        lineHeight: '1.4',
                        wordBreak: 'break-word',
                        minWidth: '280px',
                        fontSize: '0.9rem'
                      }}>{c.name}</td>

                      <td style={{ padding: '12px', textAlign: 'center', fontWeight: 'bold', color: 'var(--main-text)' }}>{c.weight}</td>
                      {projectNames.map((_, pIdx) => (
                        <td key={pIdx} style={{ padding: '8px', borderRight: '1px solid var(--main-border)', wordBreak: 'break-word' }}>
                          <select
                            value={data[pIdx][cIdx]}
                            onChange={e => handleChange(pIdx, cIdx, e.target.value)}
                            style={{
                              width: '100%',
                              padding: '10px 8px',
                              border: '1px solid var(--main-border)',
                              borderRadius: '6px',
                              fontSize: '0.81rem',
                              background: 'var(--main-surface)',
                              color: 'var(--main-text)',
                              cursor: 'pointer',
                              height: 'auto',
                              minHeight: '45px',
                              appearance: 'none',
                              backgroundImage: 'linear-gradient(45deg, transparent 50%, var(--main-text) 50%), linear-gradient(135deg, var(--main-text) 50%, transparent 50%)',
                              backgroundPosition: 'calc(100% - 15px) calc(0.7em + 10px), calc(100% - 10px) calc(0.7em + 10px)',
                              backgroundSize: '5px 5px, 5px 5px',
                              backgroundRepeat: 'no-repeat'
                            }}
                          >
                            <option value="">- Choisir -</option>
                            {c.options.map((opt: string, oIdx: number) => (
                              <option key={oIdx} value={opt} style={{ background: 'var(--main-surface)', color: 'var(--main-text)', padding: '10px' }}>
                                {opt}
                              </option>
                            ))}
                          </select>
                        </td>
                      ))}
                    </tr>
                  ))}

                  <tr style={{ background: 'rgba(0,0,0,0.08)', fontWeight: 800 }}>
                    <td colSpan={3} style={{ padding: '15px', textAlign: 'right', color: 'var(--main-text)' }}>Points totaux</td>
                    {scores.map((s, i) => (
                      <td key={i} style={{ padding: '15px', textAlign: 'center', color: s < 0 ? '#ef4444' : 'var(--main-text)', fontSize: '1rem' }}>{s}</td>
                    ))}
                  </tr>
                  <tr style={{ fontWeight: 900 }}>
                    <td colSpan={3} style={{ padding: '15px', textAlign: 'right', color: 'var(--main-text)' }}>Priorité</td>
                    {scores.map((s, i) => {
                      const prio = getPriority(s);
                      return (
                        <td key={i} style={{ padding: '15px', textAlign: 'center', backgroundColor: prio.c, color: prio.t, border: '1px solid rgba(0,0,0,0.05)' }}>
                          {prio.p}
                        </td>
                      );
                    })}
                  </tr>

                </tbody>
              </table>
            </div>
          </div>
        )}

        {viewMode === 'criteria' && (
          <div style={{ background: 'var(--main-surface)', borderRadius: '12px', padding: '1.5rem', boxShadow: '0 4px 6px -1px rgba(0,0,0,0.1)', border: '1px solid var(--main-border)' }}>
            <div style={{ overflowX: 'auto' }}>
              <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '0.8rem' }}>
                <thead style={{ background: 'var(--main-bg)', color: 'var(--table-primary)' }}>
                  <tr>
                    <th style={{ padding: '12px', border: '1px solid var(--main-border)', width: '60px' }}>Code</th>
                    <th style={{ padding: '12px', border: '1px solid var(--main-border)', textAlign: 'left', minWidth: '250px' }}>Critère</th>
                    <th style={{ padding: '12px', border: '1px solid var(--main-border)' }}>-2 (Poids min)</th>
                    <th style={{ padding: '12px', border: '1px solid var(--main-border)' }}>-1</th>
                    <th style={{ padding: '12px', border: '1px solid var(--main-border)' }}>0 (Neutre)</th>
                    <th style={{ padding: '12px', border: '1px solid var(--main-border)' }}>1</th>
                    <th style={{ padding: '12px', border: '1px solid var(--main-border)' }}>2 (Poids max)</th>
                    <th style={{ padding: '12px', border: '1px solid var(--main-border)', width: '60px' }}>Coeff.</th>
                    <th style={{ padding: '12px', border: '1px solid var(--main-border)', width: '60px' }}>Action</th>
                  </tr>
                </thead>
                <tbody>
                  {criteria.map((c: any, idx: number) => (
                    <tr key={idx} style={{ background: idx % 2 === 0 ? 'rgba(0,0,0,0.02)' : 'transparent', borderBottom: '1px solid var(--main-border)' }}>
                      <td style={{ padding: '10px', border: '1px solid var(--main-border)', textAlign: 'center', fontWeight: 'bold', color: 'var(--main-text)', background: 'rgba(0,0,0,0.05)' }}>
                        {c.code}
                      </td>
                      <td style={{ padding: '10px', border: '1px solid var(--main-border)' }}>
                        <textarea
                          value={c.name}
                          onChange={(e) => handleCriteriaChange(idx, 'name', e.target.value)}
                          rows={2}
                          style={{ width: '100%', border: 'none', background: 'transparent', color: 'var(--main-text)', resize: 'vertical', fontSize: '0.85rem', outline: 'none', fontFamily: 'inherit', fontWeight: 500 }}
                        />
                      </td>
                      {c.options.map((opt: string, oIdx: number) => (
                        <td key={oIdx} style={{ padding: '10px', border: '1px solid var(--main-border)' }}>
                          <textarea
                            value={opt}
                            onChange={(e) => {
                              const newOpts = [...c.options];
                              newOpts[oIdx] = e.target.value;
                              handleCriteriaChange(idx, 'options', newOpts);
                            }}
                            rows={2}
                            style={{ width: '100%', border: 'none', background: 'transparent', color: 'var(--main-text)', textAlign: 'center', fontSize: '0.75rem', outline: 'none', resize: 'vertical', fontFamily: 'inherit' }}
                          />
                        </td>
                      ))}
                      <td style={{ padding: '10px', border: '1px solid var(--main-border)', textAlign: 'center' }}>
                        <input
                          type="number"
                          value={c.weight}
                          onChange={(e) => handleCriteriaChange(idx, 'weight', parseInt(e.target.value) || 0)}
                          style={{ width: '40px', textAlign: 'center', border: 'none', background: 'transparent', color: 'var(--main-text)', fontWeight: 'bold', outline: 'none' }}
                        />
                      </td>
                      <td style={{ padding: '10px', border: '1px solid var(--main-border)', textAlign: 'center' }}>
                        <button
                          onClick={() => {
                            const newCriteria = criteria.filter((_: any, i: number) => i !== idx);
                            setCriteria(newCriteria);
                          }}
                          style={{ background: 'none', border: 'none', cursor: 'pointer', fontSize: '1.2rem', transition: 'transform 0.15s' }}
                          onMouseEnter={(e) => e.currentTarget.style.transform = 'scale(1.2)'}
                          onMouseLeave={(e) => e.currentTarget.style.transform = 'scale(1)'}
                          title="Supprimer ce critère"
                        >
                          🗑️
                        </button>
                      </td>
                    </tr>
                  ))}

                </tbody>
              </table>
            </div>
            <button
              onClick={() => setCriteria([...criteria, { code: `C${criteria.length + 1}`, name: 'Nouveau critère', weight: 1, options: ["-2", "-1", "0", "1", "2"] }])}
              style={{ marginTop: '1.5rem', padding: '10px 20px', background: 'var(--main-surface)', color: 'var(--table-primary)', border: '1px solid var(--table-primary)', borderRadius: '8px', cursor: 'pointer', fontWeight: 700 }}
            >
              + Ajouter un critère
            </button>
          </div>
        )}

        {viewMode === 'dashboard' && (
          <>
            <div style={{ display: 'flex', flexDirection: 'column', gap: '2.5rem', animation: 'fadeIn 0.6s ease', padding: '1.5rem', background: 'var(--main-bg)', color: 'var(--main-text)', borderRadius: '16px' }}>
              <div style={{ textAlign: 'center', marginBottom: '1rem' }}>
                <h1 style={{ fontSize: '1.8rem', color: 'var(--main-text)', marginBottom: '0.5rem', fontWeight: 700 }}>Tableau de bord de sélection des projets stratégiques</h1>
                <p style={{ fontSize: '0.9rem', color: 'var(--text-muted)', maxWidth: '800px', margin: '0 auto' }}>Récapitulatif dynamique basé sur votre <strong>Matrice d'évaluation</strong>.</p>
              </div>

              {/* Cartes de synthèse - Style Premium */}
              <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(220px, 1fr))', gap: '1.5rem' }}>
                {[
                  { title: 'Évaluations Qualité', value: projectNames.length, color: '#1e3a8a', icon: '📋' },
                  { title: 'Performance Moyenne', value: `${((scores.reduce((a, b) => a + b, 0) / (scores.length || 1) + 100) / 2).toFixed(1)}%`, color: '#991b1b', isRing: true },
                  { title: 'Impact Fort (P1)', value: scores.filter(s => s >= 50).length, color: '#065f46', icon: '🌟' },
                  { title: 'Sous la Moyenne', value: scores.filter(s => s < (scores.reduce((a, b) => a + b, 0) / (scores.length || 1))).length, color: '#9a3412', icon: '📉' }
                ].map((card, idx) => (
                  <div key={idx} style={{ background: 'var(--main-surface)', borderRadius: '12px', border: '1px solid var(--main-border)', display: 'flex', flexDirection: 'column', boxShadow: '0 4px 6px -1px rgba(0, 0, 0, 0.1)', overflow: 'hidden' }}>
                    <div style={{ background: card.color, color: '#fff', padding: '10px', fontSize: '0.8rem', fontWeight: 800, textAlign: 'center', textTransform: 'uppercase', letterSpacing: '0.5px' }}>
                      {card.title}
                    </div>
                    <div style={{ flex: 1, display: 'flex', alignItems: 'center', justifyContent: 'center', padding: '1.5rem 1rem' }}>
                      {card.isRing ? (
                        <div style={{ position: 'relative', width: '90px', height: '90px' }}>
                          <svg width="90" height="90" viewBox="0 0 90 90">
                            <circle cx="45" cy="45" r="38" fill="none" stroke="var(--main-bg)" strokeWidth="10" />
                            <circle cx="45" cy="45" r="38" fill="none" stroke={card.color} strokeWidth="10" strokeDasharray={`${parseFloat(card.value.toString()) * 2.38}, 238`} transform="rotate(-90 45 45)" strokeLinecap="round" />
                          </svg>
                          <div style={{ position: 'absolute', top: 0, left: 0, right: 0, bottom: 0, display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: '1.2rem', fontWeight: 900 }}>{card.value}</div>
                        </div>
                      ) : (
                        <div style={{ display: 'flex', alignItems: 'baseline', gap: '8px' }}>
                          <span style={{ fontSize: '1.5rem' }}>{card.icon}</span>
                          <div style={{ fontSize: '2.5rem', fontWeight: 900, color: 'var(--main-text)' }}>{card.value}</div>
                        </div>
                      )}
                    </div>
                  </div>
                ))}
              </div>

              {/* Graphiques principaux - Traduit */}
              <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(400px, 1fr))', gap: '2rem' }}>
                <div style={{ background: 'var(--main-surface)', padding: '2rem', borderRadius: '12px', border: '1px solid var(--main-border)', boxShadow: '0 4px 6px -1px rgba(0, 0, 0, 0.1)' }}>
                  <div style={{ background: '#1e3a8a', color: '#fff', padding: '8px 20px', fontSize: '0.9rem', fontWeight: 800, width: 'fit-content', borderRadius: '4px', marginBottom: '2rem' }}>
                    🏆 Classement des Projets (Points)
                  </div>
                  <div style={{ display: 'flex', flexDirection: 'column', gap: '14px' }}>
                    {projectNames.slice(0, 10).map((name, i) => {
                      const score = scores[i] || 0;
                      const normScore = score + 100;
                      const barColor = score >= 32 ? '#10b981' : score >= 0 ? '#3b82f6' : '#ef4444';
                      return (
                        <div key={i} style={{ display: 'flex', alignItems: 'center', gap: '1.2rem' }}>
                          <div style={{ width: '160px', fontSize: '0.85rem', fontWeight: 800, color: 'var(--main-text)', textAlign: 'right', whiteSpace: 'normal', wordBreak: 'break-word', lineHeight: '1.2' }} title={name}>{name}</div>

                          <div style={{ flex: 1, height: '14px', background: 'var(--main-bg)', borderRadius: '7px', position: 'relative', overflow: 'hidden', border: '1px solid var(--main-border)' }}>
                            <div style={{ width: `${(normScore / 200) * 100}%`, height: '100%', background: barColor, borderRadius: '0 4px 4px 0', transition: 'width 1s cubic-bezier(0.4, 0, 0.2, 1)' }} />
                          </div>
                          <div style={{ width: '40px', fontSize: '0.85rem', fontWeight: 800, color: barColor }}>{score}</div>
                        </div>
                      );
                    })}
                  </div>
                </div>

              </div>

              <div style={{ background: 'var(--main-surface)', padding: '2rem', borderRadius: '12px', border: '1px solid var(--main-border)', boxShadow: '0 4px 6px -1px rgba(0, 0, 0, 0.1)', position: 'relative' }}>
                <div style={{ background: '#991b1b', color: '#fff', padding: '8px 20px', fontSize: '0.9rem', fontWeight: 800, width: 'fit-content', borderRadius: '4px', marginBottom: '2rem' }}>
                  📊 Répartition des Priorités
                </div>
                <div style={{ display: 'flex', justifyContent: 'center', alignItems: 'center', height: '100%', gap: '2rem' }}>
                  <svg width="180" height="180" viewBox="0 0 42 42">
                    <circle cx="21" cy="21" r="15.9155" fill="none" stroke="var(--main-bg)" strokeWidth="6" />
                    <circle cx="21" cy="21" r="15.9155" fill="none" stroke="#10b981" strokeWidth="6" strokeDasharray="30 70" strokeDashoffset="25" />
                    <circle cx="21" cy="21" r="15.9155" fill="none" stroke="#3b82f6" strokeWidth="6" strokeDasharray="25 75" strokeDashoffset="-5" />
                    <circle cx="21" cy="21" r="15.9155" fill="none" stroke="#f59e0b" strokeWidth="6" strokeDasharray="20 80" strokeDashoffset="-30" />
                    <circle cx="21" cy="21" r="15.9155" fill="none" stroke="#ef4444" strokeWidth="6" strokeDasharray="25 75" strokeDashoffset="-50" />
                  </svg>
                  <div style={{ display: 'flex', flexDirection: 'column', gap: '10px' }}>
                    {[
                      { label: 'P1 (Élevé)', color: '#10b981' },
                      { label: 'P2 (Moyen+)', color: '#3b82f6' },
                      { label: 'P3 (Moyen-)', color: '#f59e0b' },
                      { label: 'P4 (Faible)', color: '#ef4444' }
                    ].map((p, i) => (
                      <div key={i} style={{ display: 'flex', alignItems: 'center', gap: '10px', fontSize: '0.85rem', color: 'var(--text-muted)' }}>
                        <div style={{ width: '12px', height: '12px', background: p.color, borderRadius: '2px' }} />
                        <span style={{ fontWeight: 700 }}>{p.label}</span>
                      </div>
                    ))}
                  </div>
                </div>
              </div>
              {/* Analyse Linéaire par Projet */}
              <div style={{ background: 'var(--main-surface)', padding: '2rem', borderRadius: '12px', border: '1px solid var(--main-border)', boxShadow: '0 4px 6px -1px rgba(0, 0, 0, 0.1)' }}>
                <h3 style={{ margin: '0 0 1.5rem 0', color: 'var(--main-text)', fontSize: '1.2rem', fontWeight: 800 }}>📂 Analyse Linéaire des Projets</h3>
                <div style={{ display: 'flex', flexDirection: 'column', gap: '10px' }}>
                  {projectNames.map((name, i) => {
                    const score = scores[i] || 0;
                    const statusColor = score >= 32 ? '#10b981' : score >= 0 ? '#3b82f6' : '#ef4444';
                    return (
                      <div key={i} style={{ display: 'flex', alignItems: 'center', gap: '2rem', padding: '12px 15px', borderBottom: '1px solid var(--main-border)', transition: 'background 0.2s' }}>
                        <div style={{ width: '180px', fontWeight: 700, fontSize: '0.9rem' }}>{name}</div>
                        <div style={{ flex: 1, height: '8px', background: 'var(--main-bg)', borderRadius: '4px', overflow: 'hidden' }}>
                          <div style={{ width: `${Math.min(100, Math.max(0, (score + 100) / 2))}%`, height: '100%', background: statusColor, borderRadius: '4px' }} />
                        </div>
                        <div style={{ width: '50px', textAlign: 'right', fontWeight: 900, color: statusColor }}>{score}</div>
                        <div style={{ fontSize: '0.85rem', color: 'var(--text-muted)', fontStyle: 'italic', flex: 1.5, marginLeft: '1rem', padding: '10px', background: 'rgba(0,0,0,0.02)', borderRadius: '8px' }}>
                          {score >= 32 ? "✅ Alignement stratégique fort. À prioriser." : score >= 0 ? "🆗 Potentiel modéré. Analyse de cadrage." : "⚠️ Risques élevés. Vigilance requise."}
                        </div>
                      </div>
                    );
                  })}
                </div>
              </div>
            </div>
            <div style={{ textAlign: 'center', padding: '1rem', borderTop: '1px solid var(--main-border)', fontSize: '0.75rem', color: 'var(--text-muted)', fontStyle: 'italic' }}>
              * Ce graphique cartographie les données pertinentes et change automatiquement en fonction de vos sélections dans la Matrice.
            </div>
          </>
        )}


        {viewMode === 'dossiers' && (
          <SelectionDossiersView
            key={`eval-${projectId}-${fileId}`}
            projectId={projectId}
            fileId={fileId}
            title="Dossiers d'Évaluation des Projets"
            onSelectDossier={(id, autoNav) => {
              if (onDossierChange) {
                onDossierChange(id);
              }
              if (autoNav) setViewMode('evaluation');
            }}
            onOpenDossier={(id) => {
              if (onDossierChange) {
                onDossierChange(id);
              }
              setViewMode('evaluation');
            }}
            currentDossierId={currentDossierId}
          />
        )}

        {viewMode === 'settings' && (
          <div style={{ background: 'var(--main-surface)', borderRadius: '12px', padding: '2rem', border: '1px solid var(--main-border)' }}>
            <h3>⚙️ Paramétrage de la sélection</h3>
            <p>Configurez ici les paramètres globaux pour l'évaluation des projets.</p>
            {/* Add some dummy settings for now */}
            <div style={{ marginTop: '2rem' }}>
              <label style={{ display: 'block', marginBottom: '0.5rem' }}>Seuil de priorité P1</label>
              <input type="number" defaultValue={32} style={{ padding: '8px', borderRadius: '4px', border: '1px solid var(--main-border)', background: 'transparent', color: 'inherit' }} />
            </div>
          </div>
        )}
      </div>
    </div>
  );
};



const PartiesPrenantesTableView = ({ fileId, projectId, currentDossierId, initialView, onDossierChange }: {
  fileId: string;
  projectId: string;
  currentDossierId: string | null;
  initialView?: 'registry' | 'matrix' | 'dossiers' | 'settings';
  onDossierChange?: (id: string) => void;
}) => {
  const [viewMode, setViewMode] = useState<'registry' | 'matrix' | 'dossiers' | 'settings'>(initialView || 'registry');

  useEffect(() => {
    if (initialView) setViewMode(initialView);
  }, [initialView]);

  const dossierSubKey = currentDossierId ? `-${currentDossierId}` : "";
  const storageKey = `registre-parties-variante1-${projectId}-${fileId}${dossierSubKey}`;

  const defaultStakeholders = [
    { id: 'P1', nom: 'Château Daniel', type: 'Interne', role: 'Directeur Informatique', interet: 'Faible', pouvoir: 'Faible', strategie: 'Effort Minimal', contributions: 'Allocation de ressources', attentes: 'Périmètre et délais', actions: 'Inviter au codir, Partager le planning...', contact: 'ChateauD@confassis.com' },
    { id: 'P2', nom: 'Meunier Jacques', type: 'Interne', role: 'Responsable administratif', interet: 'Élevé', pouvoir: 'Faible', strategie: 'Garder Informé', contributions: 'Intégration et Formation', attentes: 'Périmètre et délais', actions: 'Inviter au codir, Partager la matrice...', contact: 'MeunierJ@confassis.com' },
    { id: 'P3', nom: 'Robespierre Olivier', type: 'Interne', role: 'Directeur Financier CMRL', interet: 'Faible', pouvoir: 'Élevé', strategie: 'Garder Satisfait', contributions: 'Sponsor', attentes: 'Budget', actions: 'Inviter au codir, Partager le suivi...', contact: 'RobespierreO@confassis.com' }
  ];

  const [stakeholders, setStakeholders] = useState(() => {
    const saved = localStorage.getItem(storageKey);
    return saved ? JSON.parse(saved) : defaultStakeholders;
  });

  useEffect(() => {
    localStorage.setItem(storageKey, JSON.stringify(stakeholders));
  }, [stakeholders, storageKey]);

  // Sync when dossier changes
  useEffect(() => {
    const saved = localStorage.getItem(storageKey);
    if (saved) setStakeholders(JSON.parse(saved));
    else setStakeholders(defaultStakeholders);
  }, [storageKey]);

  const updateStrategy = (pouvoir: string, interet: string) => {
    if (pouvoir === 'Élevé') {
      return interet === 'Élevé' ? 'Acteur Clé' : 'Garder Satisfait';
    } else {
      return interet === 'Élevé' ? 'Garder Informé' : 'Effort Minimal';
    }
  };

  const handleChange = (id: string, field: string, val: any) => {
    setStakeholders(stakeholders.map((s: any) => {
      if (s.id === id) {
        const updated = { ...s, [field]: val };
        if (field === 'pouvoir' || field === 'interet') {
          updated.strategie = updateStrategy(updated.pouvoir, updated.interet);
        }
        return updated;
      }
      return s;
    }));
  };

  return (
    <div style={{ flex: 1, display: 'flex', flexDirection: 'column', backgroundColor: 'var(--main-bg)', color: 'var(--main-text)', overflow: 'hidden' }}>
      <div style={{ flex: 1, overflow: 'auto', padding: '1.5rem' }}>

        {viewMode === 'registry' && (
          <div style={{ background: 'var(--main-surface)', borderRadius: '12px', padding: '1.5rem', boxShadow: '0 4px 6px -1px rgba(0,0,0,0.1)', border: '1px solid var(--main-border)' }}>
            <h3 style={{ margin: '0 0 1.5rem 0' }}>📋 Registre des parties prenantes</h3>
            <div style={{ overflowX: 'auto' }}>
              <table style={{ minWidth: '1200px', width: '100%', borderCollapse: 'collapse', fontSize: '0.85rem' }}>
                <thead style={{ backgroundColor: 'var(--table-primary)', color: '#fff' }}>
                  <tr>
                    <th style={{ padding: '12px', width: '60px', borderRight: '1px solid rgba(255,255,255,0.1)' }}>ID</th>
                    <th style={{ padding: '12px', textAlign: 'left', borderRight: '1px solid rgba(255,255,255,0.1)' }}>Nom</th>
                    <th style={{ padding: '12px', width: '100px', borderRight: '1px solid rgba(255,255,255,0.1)' }}>Type</th>
                    <th style={{ padding: '12px', borderRight: '1px solid rgba(255,255,255,0.1)' }}>Rôle</th>
                    <th style={{ padding: '12px', width: '100px', borderRight: '1px solid rgba(255,255,255,0.1)' }}>Intérêt</th>
                    <th style={{ padding: '12px', width: '100px', borderRight: '1px solid rgba(255,255,255,0.1)' }}>Pouvoir</th>
                    <th style={{ padding: '12px', width: '120px', borderRight: '1px solid rgba(255,255,255,0.1)' }}>Stratégie</th>
                    <th style={{ padding: '12px', borderRight: '1px solid rgba(255,255,255,0.1)' }}>Contributions</th>
                    <th style={{ padding: '12px', borderRight: '1px solid rgba(255,255,255,0.1)' }}>Attentes</th>
                    <th style={{ padding: '12px', width: '60px' }}>Action</th>
                  </tr>
                </thead>
                <tbody>
                  {stakeholders.map((s: any, idx: number) => (
                    <tr key={s.id} style={{ borderBottom: '1px solid var(--main-border)', background: idx % 2 === 0 ? 'rgba(0,0,0,0.01)' : 'transparent' }}>
                      <td style={{ padding: '10px', textAlign: 'center', backgroundColor: 'var(--main-bg)', fontWeight: 'bold' }}>{s.id}</td>
                      <td style={{ padding: '8px' }}>
                        <textarea
                          value={s.nom}
                          onChange={(e) => handleChange(s.id, 'nom', e.target.value)}
                          rows={2}
                          style={{ width: '100%', border: 'none', background: 'rgba(0,0,0,0.02)', padding: '6px', borderRadius: '4px', color: 'inherit', fontWeight: 600, resize: 'none' }}
                        />
                      </td>
                      <td style={{ padding: '8px' }}>
                        <select value={s.type} onChange={(e) => handleChange(s.id, 'type', e.target.value)} style={{ width: '100%', padding: '6px', borderRadius: '4px', border: '1px solid var(--main-border)', background: 'var(--main-surface)', color: 'inherit' }}>
                          <option value="Interne">Interne</option>
                          <option value="Externe">Externe</option>
                        </select>
                      </td>
                      <td style={{ padding: '8px' }}>
                        <textarea value={s.role} onChange={(e) => handleChange(s.id, 'role', e.target.value)} rows={2} style={{ width: '100%', border: 'none', background: 'rgba(0,0,0,0.02)', padding: '6px', borderRadius: '4px', color: 'inherit', fontSize: '0.8rem', resize: 'none' }} />
                      </td>
                      <td style={{ padding: '8px' }}>
                        <select value={s.interet} onChange={(e) => handleChange(s.id, 'interet', e.target.value)} style={{ width: '100%', padding: '6px', borderRadius: '4px', border: '1px solid var(--main-border)', background: 'var(--main-surface)', color: 'inherit' }}>
                          <option value="Faible">Faible</option>
                          <option value="Élevé">Élevé</option>
                        </select>
                      </td>
                      <td style={{ padding: '8px' }}>
                        <select value={s.pouvoir} onChange={(e) => handleChange(s.id, 'pouvoir', e.target.value)} style={{ width: '100%', padding: '6px', borderRadius: '4px', border: '1px solid var(--main-border)', background: 'var(--main-surface)', color: 'inherit' }}>
                          <option value="Faible">Faible</option>
                          <option value="Élevé">Élevé</option>
                        </select>
                      </td>
                      <td style={{ padding: '10px', textAlign: 'center' }}>
                        <span style={{
                          padding: '4px 10px',
                          borderRadius: '12px',
                          fontSize: '0.7rem',
                          fontWeight: 800,
                          whiteSpace: 'nowrap',
                          backgroundColor: s.strategie === 'Acteur Clé' ? '#ef4444' : s.strategie === 'Garder Satisfait' ? '#f59e0b' : s.strategie === 'Garder Informé' ? '#3b82f6' : '#10b981',
                          color: '#fff'
                        }}>
                          {s.strategie}
                        </span>
                      </td>
                      <td style={{ padding: '8px' }}>
                        <textarea value={s.contributions} onChange={(e) => handleChange(s.id, 'contributions', e.target.value)} rows={2} style={{ width: '100%', border: 'none', background: 'rgba(0,0,0,0.02)', padding: '6px', borderRadius: '4px', color: 'inherit', fontSize: '0.8rem', resize: 'none' }} />
                      </td>
                      <td style={{ padding: '8px' }}>
                        <textarea value={s.attentes} onChange={(e) => handleChange(s.id, 'attentes', e.target.value)} rows={2} style={{ width: '100%', border: 'none', background: 'rgba(0,0,0,0.02)', padding: '6px', borderRadius: '4px', color: 'inherit', fontSize: '0.8rem', resize: 'none' }} />
                      </td>
                      <td style={{ padding: '8px', textAlign: 'center' }}>
                        <button onClick={() => setStakeholders(stakeholders.filter((it: any) => it.id !== s.id))} style={{ background: 'none', border: 'none', cursor: 'pointer', fontSize: '1.1rem' }}>🗑️</button>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
            <button
              onClick={() => setStakeholders([...stakeholders, { id: `P${Date.now().toString().slice(-4)}`, nom: '', type: 'Interne', role: '', interet: 'Faible', pouvoir: 'Faible', strategie: 'Effort Minimal', contributions: '', attentes: '', actions: '', contact: '' }])}
              style={{ marginTop: '1.5rem', padding: '10px 24px', background: 'var(--table-primary)', color: '#fff', border: 'none', borderRadius: '8px', cursor: 'pointer', fontWeight: 800 }}
            >
              + Ajouter une partie prenante
            </button>
          </div>
        )}

        {viewMode === 'matrix' && (
          <PartiesPrenantesMapView stakeholders={stakeholders} />
        )}

        {viewMode === 'dossiers' && (
          <SelectionDossiersView
            key={`prenantes-${projectId}-${fileId}`}
            projectId={projectId}
            fileId={fileId}
            title="Dossiers des Parties Prenantes"
            onSelectDossier={(id, autoNav) => {
              onDossierChange?.(id);
              if (autoNav) setViewMode('registry');
            }}
            onOpenDossier={(id) => {
              onDossierChange?.(id);
              setViewMode('registry');
            }}
            currentDossierId={currentDossierId}
          />
        )}

        {viewMode === 'settings' && (
          <PartiesPrenantesSettingsView />
        )}

      </div>
    </div>
  );
};

const FaisabiliteMatriceView = ({ fileId, projectId, currentDossierId, initialView, onDossierChange }: {
  fileId: string;
  projectId: string;
  currentDossierId: string | null;
  initialView?: 'evaluation' | 'criteria' | 'dashboard' | 'database' | 'dossiers' | 'settings';
  onDossierChange?: (id: string) => void;
}) => {
  const [viewMode, setViewMode] = useState<'evaluation' | 'criteria' | 'dashboard' | 'database' | 'dossiers' | 'settings'>(initialView || 'evaluation');

  useEffect(() => {
    if (initialView) setViewMode(initialView);
  }, [initialView]);

  const dossierSubKey = currentDossierId ? `-${currentDossierId}` : "";
  const storageKey = `matrice-faisabilite-${projectId}-${fileId}${dossierSubKey}`;

  const [criteria, setCriteria] = useState(() => {
    const saved = localStorage.getItem(`faisabilite-criteria-${projectId}`);
    if (saved) return JSON.parse(saved);
    return [
      { code: 'C1', section: 'Commerciale', name: 'Le prix du produit est attractif', weight: 3 },
      { code: 'C2', section: 'Commerciale', name: 'Le délai de développement est cohérent avec le marché', weight: 4 },
      { code: 'C3', section: 'Commerciale', name: 'Les ventes souhaitées semblent atteignables', weight: 5 },
      { code: 'C4', section: 'Technique', name: 'Le projet s\'inscrit dans la stratégie technique', weight: 3 },
      { code: 'C5', section: 'Technique', name: 'Le produit répond à une attente technologique', weight: 4 },
      { code: 'C6', section: 'Technique', name: 'Faisabilité technique interne assurée', weight: 2 },
      { code: 'C7', section: 'Financière', name: 'Rentabilité prévisionnelle acceptable', weight: 5 },
      { code: 'C8', section: 'Financière', name: 'Disponibilité des financements', weight: 4 },
      { code: 'C9', section: 'Financière', name: 'Impact sur le cash-flow positif à court terme', weight: 3 },
      { code: 'C10', section: 'Economique/Sociale', name: 'Impact positif sur l\'image de marque', weight: 3 },
      { code: 'C11', section: 'Economique/Sociale', name: 'Acceptation sociale et environnementale', weight: 2 }
    ];
  });

  const [projectNames, setProjectNames] = useState<string[]>(() => {
    const saved = localStorage.getItem(`${storageKey}-names`);
    if (saved) return JSON.parse(saved);
    return ["Projet Alpha", "Projet Beta", "Projet Gamma"];
  });

  const choices = ["Pas du tout d'accord", "Plutôt pas d'accord", "Plutôt d'accord", "Tout à fait d'accord"];

  const getChoiceValue = (val: string) => {
    const idx = choices.indexOf(val);
    return idx !== -1 ? idx + 1 : 0;
  };

  const [data, setData] = useState<string[][]>(() => {
    const saved = localStorage.getItem(storageKey);
    if (saved) return JSON.parse(saved);
    return Array(3).fill(null).map(() =>
      Array(criteria.length).fill('').map(() => choices[Math.floor(Math.random() * choices.length)])
    );
  });

  useEffect(() => {
    localStorage.setItem(storageKey, JSON.stringify(data));
  }, [data, storageKey]);

  useEffect(() => {
    localStorage.setItem(`${storageKey}-names`, JSON.stringify(projectNames));
  }, [projectNames, storageKey]);

  useEffect(() => {
    localStorage.setItem(`faisabilite-criteria-${projectId}`, JSON.stringify(criteria));
  }, [criteria, projectId]);

  // Sync data when storageKey changes
  useEffect(() => {
    const savedNames = localStorage.getItem(`${storageKey}-names`);
    if (savedNames) setProjectNames(JSON.parse(savedNames));
    const savedData = localStorage.getItem(storageKey);
    if (savedData) setData(JSON.parse(savedData));
  }, [storageKey]);

  const handleProjectNameChange = (idx: number, name: string) => {
    const newNames = [...projectNames];
    newNames[idx] = name;
    setProjectNames(newNames);
  };

  const handleChange = (pIdx: number, cIdx: number, val: string) => {
    const newData = [...data];
    if (!newData[pIdx]) newData[pIdx] = Array(criteria.length).fill('');
    newData[pIdx][cIdx] = val;
    setData(newData);
  };

  const handleCriteriaChange = (idx: number, field: string, val: any) => {
    const newCriteria = [...criteria];
    newCriteria[idx] = { ...newCriteria[idx], [field]: val };
    setCriteria(newCriteria);
  };

  const scores = projectNames.map((_, pIdx) => {
    let totalWeight = 0;
    let totalScore = 0;
    criteria.forEach((c: any, cIdx: number) => {
      const val = data[pIdx]?.[cIdx];
      const v = getChoiceValue(val);
      if (v > 0) {
        totalWeight += c.weight;
        totalScore += (v * c.weight);
      }
    });
    return totalWeight > 0 ? (totalScore / totalWeight) : 0;
  });

  const getAppreciation = (score: number) => {
    if (score === 0) return { label: 'A évaluer', color: '#94a3b8' };
    if (score <= 2.0) return { label: 'NON SATISFAISANTE', color: '#ef4444' };
    if (score <= 3.0) return { label: 'PEU SATISFAISANTE', color: '#f97316' };
    if (score <= 3.5) return { label: 'SATISFAISANTE', color: '#84cc16' };
    return { label: 'TRÈS SATISFAISANTE', color: '#22c55e' };
  };

  // Group criteria by sections
  const sections = Array.from(new Set(criteria.map((c: any) => c.section)));

  return (
    <div style={{ flex: 1, display: 'flex', flexDirection: 'column', backgroundColor: 'var(--main-bg)', color: 'var(--main-text)', overflow: 'hidden' }}>
      <div style={{ flex: 1, overflow: 'auto', padding: '1.5rem' }}>

        {viewMode === 'database' && (
          <div style={{ background: 'var(--main-surface)', borderRadius: '12px', padding: '1.5rem', boxShadow: '0 4px 6px -1px rgba(0,0,0,0.1)', border: '1px solid var(--main-border)' }}>
            <h3 style={{ margin: '0 0 1.5rem 0' }}>📋 Liste des projets à évaluer</h3>
            <table style={{ width: '100%', borderCollapse: 'collapse', borderRadius: '8px', overflow: 'hidden' }}>
              <thead style={{ background: 'var(--table-primary)', color: '#fff' }}>
                <tr>
                  <th style={{ padding: '12px', textAlign: 'left', width: '60px' }}>#</th>
                  <th style={{ padding: '12px', textAlign: 'left' }}>Nom du Projet</th>
                  <th style={{ padding: '12px', textAlign: 'center', width: '120px' }}>Actions</th>
                </tr>
              </thead>
              <tbody>
                {projectNames.map((p, i) => (
                  <tr key={i} style={{ borderBottom: '1px solid var(--main-border)', background: i % 2 === 0 ? 'rgba(0,0,0,0.02)' : 'transparent' }}>
                    <td style={{ padding: '12px', fontWeight: 700 }}>{i + 1}</td>
                    <td style={{ padding: '12px' }}>
                      <input
                        value={p}
                        onChange={(e) => handleProjectNameChange(i, e.target.value)}
                        style={{ width: '100%', padding: '8px', borderRadius: '4px', border: '1px solid var(--main-border)', background: 'rgba(0,0,0,0.02)', color: 'var(--main-text)', fontWeight: 600 }}
                      />
                    </td>
                    <td style={{ padding: '12px', textAlign: 'center' }}>
                      <button
                        onClick={() => {
                          const newNames = projectNames.filter((_, idx) => idx !== i);
                          const newData = data.filter((_, idx) => idx !== i);
                          setProjectNames(newNames);
                          setData(newData);
                        }}
                        style={{ background: 'none', border: 'none', cursor: 'pointer', fontSize: '1.2rem', transition: 'transform 0.15s' }}
                        onMouseEnter={(e) => e.currentTarget.style.transform = 'scale(1.2)'}
                        onMouseLeave={(e) => e.currentTarget.style.transform = 'scale(1)'}
                        title="Supprimer ce projet"
                      >
                        🗑️
                      </button>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
            <button
              onClick={() => {
                setProjectNames([...projectNames, `Nouveau Projet ${projectNames.length + 1}`]);
                setData([...data, Array(criteria.length).fill('')]);
              }}
              style={{ marginTop: '1.5rem', padding: '10px 24px', background: 'var(--table-primary)', color: '#fff', border: 'none', borderRadius: '8px', cursor: 'pointer', fontWeight: 800, fontSize: '0.9rem', boxShadow: '0 4px 12px rgba(59, 130, 246, 0.3)' }}
            >
              + Ajouter un projet
            </button>
          </div>
        )}

        {viewMode === 'evaluation' && (
          <div style={{ background: 'var(--main-surface)', borderRadius: '12px', padding: '1.5rem', boxShadow: '0 4px 6px -1px rgba(0,0,0,0.1)', border: '1px solid var(--main-border)' }}>
            <div style={{ overflowX: 'auto' }}>
              <div style={{ overflowX: 'auto' }}>
                <table style={{ minWidth: '1000px', width: '100%', borderCollapse: 'collapse', fontSize: '0.85rem' }}>
                  <thead>
                    <tr style={{ backgroundColor: 'var(--table-primary)', color: '#fff' }}>
                      <th style={{ padding: '12px', width: '60px', border: '1px solid rgba(255,255,255,0.1)' }}>Code</th>
                      <th style={{ padding: '12px', textAlign: 'left', border: '1px solid rgba(255,255,255,0.1)' }}>Critères de faisabilité</th>
                      <th style={{ padding: '12px', width: '80px', border: '1px solid rgba(255,255,255,0.1)' }}>Coeff.</th>
                      {projectNames.map((p, i) => (
                        <th key={i} style={{ padding: '12px', minWidth: '220px', border: '1px solid rgba(255,255,255,0.1)' }}>
                          <textarea
                            value={p}
                            onChange={(e) => handleProjectNameChange(i, e.target.value)}
                            rows={2}
                            style={{
                              width: '100%',
                              border: 'none',
                              background: 'rgba(255,255,255,0.15)',
                              color: '#fff',
                              textAlign: 'center',
                              fontSize: '0.85rem',
                              fontWeight: 800,
                              padding: '8px',
                              borderRadius: '6px',
                              outline: 'none',
                              resize: 'none',
                              fontFamily: 'inherit'
                            }}
                          />
                        </th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {sections.map((sectionName: any) => (
                      <Fragment key={String(sectionName)}>
                        <tr style={{ background: 'rgba(59, 130, 246, 0.1)' }}>
                          <td colSpan={3 + projectNames.length} style={{ padding: '10px 15px', fontWeight: 900, color: 'var(--table-primary)', fontSize: '0.9rem', borderLeft: '4px solid var(--table-primary)' }}>
                            {sectionName ? sectionName.toUpperCase() : 'AUTRES'}
                          </td>
                        </tr>
                        {criteria.filter((c: any) => c.section === sectionName).map((c: any) => {
                          const globalCIdx = criteria.findIndex((cr: any) => cr === c);
                          return (
                            <tr key={globalCIdx} style={{ borderBottom: '1px solid var(--main-border)' }}>
                              <td style={{ padding: '12px', textAlign: 'center', fontWeight: 'bold' }}>{c.code}</td>
                              <td style={{ padding: '12px', fontWeight: 500 }}>{c.name}</td>
                              <td style={{ padding: '12px', textAlign: 'center', fontWeight: 'bold' }}>{c.weight}</td>
                              {projectNames.map((_, pIdx) => (
                                <td key={pIdx} style={{ padding: '8px' }}>
                                  <select
                                    value={data[pIdx]?.[globalCIdx] || ''}
                                    onChange={e => handleChange(pIdx, globalCIdx, e.target.value)}
                                    style={{
                                      width: '100%',
                                      padding: '8px',
                                      borderRadius: '6px',
                                      border: '1px solid var(--main-border)',
                                      background: 'var(--main-surface)',
                                      color: 'var(--main-text)',
                                      fontSize: '0.8rem'
                                    }}
                                  >
                                    <option value="">- Choisir -</option>
                                    {choices.map(opt => <option key={opt} value={opt}>{opt}</option>)}
                                  </select>
                                </td>
                              ))}
                            </tr>
                          );
                        })}
                      </Fragment>
                    ))}
                    <tr style={{ background: 'rgba(0,0,0,0.08)', fontWeight: 800 }}>
                      <td colSpan={3} style={{ padding: '15px', textAlign: 'right' }}>Score moyen ( / 4)</td>
                      {scores.map((s, i) => (
                        <td key={i} style={{ padding: '15px', textAlign: 'center', fontSize: '1rem' }}>{s.toFixed(2)}</td>
                      ))}
                    </tr>
                    <tr style={{ fontWeight: 900 }}>
                      <td colSpan={3} style={{ padding: '15px', textAlign: 'right' }}>Appréciation</td>
                      {scores.map((s, i) => {
                        const app = getAppreciation(s);
                        return (
                          <td key={i} style={{ padding: '15px', textAlign: 'center', backgroundColor: app.color, color: '#fff' }}>
                            {app.label}
                          </td>
                        );
                      })}
                    </tr>
                  </tbody>
                </table>
              </div>
            </div>
          </div>
        )}

        {viewMode === 'criteria' && (
          <div style={{ background: 'var(--main-surface)', borderRadius: '12px', padding: '1.5rem', boxShadow: '0 4px 6px -1px rgba(0,0,0,0.1)', border: '1px solid var(--main-border)' }}>
            <h3 style={{ margin: '0 0 1.5rem 0' }}>📑 Référentiel des critères de faisabilité</h3>
            <table style={{ width: '100%', borderCollapse: 'collapse' }}>
              <thead style={{ background: 'var(--main-bg)' }}>
                <tr>
                  <th style={{ padding: '12px', border: '1px solid var(--main-border)', width: '80px' }}>Code</th>
                  <th style={{ padding: '12px', border: '1px solid var(--main-border)', width: '150px' }}>Section</th>
                  <th style={{ padding: '12px', border: '1px solid var(--main-border)', textAlign: 'left' }}>Critère</th>
                  <th style={{ padding: '12px', border: '1px solid var(--main-border)', width: '80px' }}>Coeff.</th>
                  <th style={{ padding: '12px', border: '1px solid var(--main-border)', width: '100px' }}>Action</th>
                </tr>
              </thead>
              <tbody>
                {criteria.map((c: any, idx: number) => (
                  <tr key={idx}>
                    <td style={{ padding: '10px', border: '1px solid var(--main-border)', textAlign: 'center' }}>
                      <input value={c.code} onChange={e => handleCriteriaChange(idx, 'code', e.target.value)} style={{ width: '100%', border: 'none', background: 'transparent', textAlign: 'center', color: 'inherit' }} />
                    </td>
                    <td style={{ padding: '10px', border: '1px solid var(--main-border)' }}>
                      <input value={c.section} onChange={e => handleCriteriaChange(idx, 'section', e.target.value)} style={{ width: '100%', border: 'none', background: 'transparent', color: 'inherit' }} />
                    </td>
                    <td style={{ padding: '10px', border: '1px solid var(--main-border)' }}>
                      <textarea value={c.name} onChange={e => handleCriteriaChange(idx, 'name', e.target.value)} rows={2} style={{ width: '100%', border: 'none', background: 'transparent', color: 'inherit', resize: 'vertical' }} />
                    </td>
                    <td style={{ padding: '10px', border: '1px solid var(--main-border)', textAlign: 'center' }}>
                      <input type="number" value={c.weight} onChange={e => handleCriteriaChange(idx, 'weight', parseInt(e.target.value) || 0)} style={{ width: '100%', border: 'none', background: 'transparent', textAlign: 'center', color: 'inherit' }} />
                    </td>
                    <td style={{ padding: '10px', border: '1px solid var(--main-border)', textAlign: 'center' }}>
                      <button onClick={() => setCriteria(criteria.filter((_: any, i: number) => i !== idx))} style={{ background: 'none', border: 'none', cursor: 'pointer' }}>🗑️</button>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
            <button
              onClick={() => setCriteria([...criteria, { code: `C${criteria.length + 1}`, section: 'Nouvelle Section', name: 'Nouveau critère', weight: 1 }])}
              style={{ marginTop: '1.5rem', padding: '10px 20px', background: 'var(--table-primary)', color: '#fff', border: 'none', borderRadius: '8px', cursor: 'pointer', fontWeight: 700 }}
            >
              + Ajouter un critère
            </button>
          </div>
        )}

        {viewMode === 'dashboard' && (
          <div style={{ display: 'flex', flexDirection: 'column', gap: '2rem' }}>
            <div style={{ textAlign: 'center' }}>
              <h2 style={{ fontSize: '1.8rem', fontWeight: 700 }}>Tableau de bord de faisabilité</h2>
              <p style={{ color: 'var(--text-muted)' }}>Moyennes par axe et par projet</p>
            </div>

            <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(350px, 1fr))', gap: '2rem' }}>
              {projectNames.map((name, pIdx) => {
                const pScore = scores[pIdx];
                const app = getAppreciation(pScore);
                return (
                  <div key={pIdx} style={{ background: 'var(--main-surface)', padding: '1.5rem', borderRadius: '16px', border: '1px solid var(--main-border)' }}>
                    <h3 style={{ margin: '0 0 1rem 0', color: 'var(--table-primary)' }}>{name}</h3>
                    <div style={{ display: 'flex', alignItems: 'center', gap: '1rem', marginBottom: '1.5rem' }}>
                      <div style={{ fontSize: '2.5rem', fontWeight: 900 }}>{pScore.toFixed(2)}</div>
                      <div style={{ background: app.color, color: '#fff', padding: '4px 12px', borderRadius: '20px', fontSize: '0.8rem', fontWeight: 800 }}>{app.label}</div>
                    </div>
                    {sections.map(sec => {
                      let sWeight = 0;
                      let sScore = 0;
                      criteria.forEach((c: any, cIdx: number) => {
                        if (c.section === sec) {
                          const val = data[pIdx][cIdx];
                          const v = getChoiceValue(val);
                          if (v > 0) { sWeight += c.weight; sScore += (v * c.weight); }
                        }
                      });
                      const sAvg = sWeight > 0 ? (sScore / sWeight) : 0;
                      const sPct = (sAvg / 4) * 100;
                      return (
                        <div key={String(sec)} style={{ marginBottom: '1rem' }}>
                          <div style={{ display: 'flex', justifyContent: 'space-between', fontSize: '0.8rem', marginBottom: '4px' }}>
                            <span>{sec}</span>
                            <span>{sAvg.toFixed(2)} / 4</span>
                          </div>
                          <div style={{ height: '8px', background: 'rgba(0,0,0,0.1)', borderRadius: '4px', overflow: 'hidden' }}>
                            <div style={{ width: `${sPct}%`, height: '100%', background: 'var(--table-primary)' }} />
                          </div>
                        </div>
                      );
                    })}
                  </div>
                );
              })}
            </div>
          </div>
        )}

        {viewMode === 'dossiers' && (
          <SelectionDossiersView
            key={`fais-${projectId}-${fileId}`}
            projectId={projectId}
            fileId={fileId}
            title="Dossiers de Faisabilité"
            onSelectDossier={(id, autoNav) => {
              onDossierChange?.(id);
              if (autoNav) setViewMode('matrix');
            }}
            onOpenDossier={(id) => {
              onDossierChange?.(id);
              setViewMode('matrix');
            }}
            currentDossierId={currentDossierId}
          />
        )}

        {viewMode === 'settings' && (
          <div style={{ background: 'var(--main-surface)', borderRadius: '12px', padding: '1.5rem', border: '1px solid var(--main-border)' }}>
            <h3>⚙️ Paramétrage de la faisabilité</h3>
            <p>Configurez ici les options globales pour l'analyse de faisabilité.</p>
            <div style={{ marginTop: '2rem', display: 'flex', flexDirection: 'column', gap: '1rem' }}>
              <label>Seuil SATISFAISANT : <input type="number" step="0.1" defaultValue={3.0} style={{ padding: '8px', borderRadius: '4px', border: '1px solid var(--main-border)', background: 'transparent', color: 'inherit' }} /></label>
              <label>Seuil TRÈS SATISFAISANT : <input type="number" step="0.1" defaultValue={3.5} style={{ padding: '8px', borderRadius: '4px', border: '1px solid var(--main-border)', background: 'transparent', color: 'inherit' }} /></label>
            </div>
          </div>
        )}
      </div>
    </div>
  );
};


const PartiesPrenantesMapView = ({ stakeholders }: { stakeholders: any[] }) => {
  const matrixBlocks = [
    { title: 'Garder Satisfait', color: 'hsla(82, 63%, 50%, 0.8)', bg: 'hsla(82, 63%, 95%, 1)', border: 'hsla(82, 63%, 40%, 1)', p: 'Élevé', i: 'Faible', desc: 'Haut Pouvoir / Bas Intérêt' },
    { title: 'Acteur Clé', color: 'hsla(45, 98%, 51%, 0.8)', bg: 'hsla(45, 98%, 95%, 1)', border: 'hsla(45, 98%, 40%, 1)', p: 'Élevé', i: 'Élevé', desc: 'Haut Pouvoir / Haut Intérêt' },
    { title: 'Effort Minimal', color: 'hsla(217, 91%, 60%, 0.8)', bg: 'hsla(217, 91%, 95%, 1)', border: 'hsla(217, 91%, 40%, 1)', p: 'Faible', i: 'Faible', desc: 'Bas Pouvoir / Bas Intérêt' },
    { title: 'Garder Informé', color: 'hsla(25, 95%, 53%, 0.8)', bg: 'hsla(25, 95%, 95%, 1)', border: 'hsla(25, 95%, 40%, 1)', p: 'Faible', i: 'Élevé', desc: 'Bas Pouvoir / Haut Intérêt' }
  ];

  return (
    <div style={{ flex: 1, overflow: 'auto', background: 'var(--main-bg)', color: 'var(--main-text)', padding: '2rem', borderRadius: '16px' }}>
      <div style={{
        background: 'var(--main-surface)',
        color: 'var(--main-text)',
        padding: '1.25rem 2rem',
        borderRadius: '16px',
        fontSize: '1.4rem',
        fontWeight: 900,
        marginBottom: '2rem',
        boxShadow: '0 8px 32px rgba(0,0,0,0.1)',
        border: '1px solid var(--main-border)',
        display: 'flex',
        alignItems: 'center',
        gap: '1rem'
      }}>
        <span style={{ fontSize: '1.8rem' }}>🌍</span> Matrice Pouvoir-Intérêt
      </div>

      <div style={{ display: 'grid', gridTemplateColumns: 'minmax(600px, 1.5fr) 1fr', gap: '2.5rem', alignItems: 'flex-start' }}>
        <div style={{
          display: 'grid',
          gridTemplateColumns: '50px 1fr 1fr',
          gridTemplateRows: '1fr 1fr 50px',
          width: '100%',
          aspectRatio: '1.2 / 1',
          position: 'relative',
          background: 'rgba(255,255,255,0.03)',
          borderRadius: '24px',
          padding: '10px',
          border: '1px solid var(--main-border)'
        }}>
          {/* Y Axis Label */}
          <div style={{ gridRow: '1 / 3', display: 'flex', alignItems: 'center', justifyContent: 'center' }}>
            <div style={{ transform: 'rotate(-90deg)', fontWeight: 800, color: 'var(--text-muted)', width: '300px', textAlign: 'center', fontSize: '0.9rem', letterSpacing: '2px', textTransform: 'uppercase' }}>
              Pouvoir (Élevé → Faible)
            </div>
          </div>

          {matrixBlocks.map((block, idx) => (
            <div key={idx} style={{
              background: 'var(--main-surface)',
              border: `2px solid ${block.border}`,
              borderRadius: '20px',
              margin: '6px',
              padding: '1.25rem',
              display: 'flex',
              flexDirection: 'column',
              gap: '0.75rem',
              boxShadow: '0 4px 15px rgba(0,0,0,0.05)',
              position: 'relative',
              overflow: 'hidden',
              transition: 'transform 0.3s ease'
            }}
              onMouseEnter={(e) => e.currentTarget.style.transform = 'scale(1.02)'}
              onMouseLeave={(e) => e.currentTarget.style.transform = 'scale(1)'}
            >
              <div style={{
                position: 'absolute',
                top: '-10px',
                right: '-10px',
                width: '40px',
                height: '40px',
                background: block.color,
                opacity: 0.1,
                borderRadius: '50%'
              }} />

              <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', borderBottom: `1px solid ${block.color}33`, paddingBottom: '8px' }}>
                <div style={{ fontWeight: 800, color: block.border, fontSize: '1rem' }}>{block.title}</div>
                <div style={{
                  background: block.color,
                  color: '#fff',
                  width: '24px',
                  height: '24px',
                  borderRadius: '50%',
                  display: 'flex',
                  alignItems: 'center',
                  justifyContent: 'center',
                  fontSize: '0.75rem',
                  fontWeight: 900
                }}>
                  {stakeholders.filter((s: any) => s.pouvoir === block.p && s.interet === block.i).length}
                </div>
              </div>
              <div style={{ fontSize: '0.7rem', color: 'var(--text-muted)', fontStyle: 'italic', marginBottom: '4px' }}>{block.desc}</div>
              <div style={{ flex: 1, fontSize: '0.85rem', overflowY: 'auto', paddingRight: '4px' }}>
                <div style={{ display: 'flex', flexWrap: 'wrap', gap: '0.5rem' }}>
                  {stakeholders.filter((s: any) => s.pouvoir === block.p && s.interet === block.i).map((s: any) => (
                    <div key={s.id} style={{
                      background: `${block.color}15`,
                      border: `1px solid ${block.color}33`,
                      padding: '4px 10px',
                      borderRadius: '8px',
                      display: 'flex',
                      alignItems: 'center',
                      gap: '6px',
                      fontSize: '0.75rem',
                      fontWeight: 600,
                      color: 'var(--main-text)'
                    }}>
                      <span style={{ color: block.border }}>👤</span> {s.nom || `Inconnu (${s.id})`}
                    </div>
                  ))}
                  {stakeholders.filter((s: any) => s.pouvoir === block.p && s.interet === block.i).length === 0 && (
                    <div style={{ color: 'var(--text-muted)', fontSize: '0.75rem', opacity: 0.5 }}>Aucun acteur</div>
                  )}
                </div>
              </div>
            </div>
          ))}

          {/* X Axis Label */}
          <div style={{ gridColumn: '2 / 4', display: 'flex', alignItems: 'center', justifyContent: 'center', paddingTop: '10px' }}>
            <div style={{ fontWeight: 800, color: 'var(--text-muted)', fontSize: '0.9rem', letterSpacing: '2px', textTransform: 'uppercase' }}>
              Intérêt (Faible → Élevé)
            </div>
          </div>
        </div>

        <div style={{
          background: 'var(--main-surface)',
          padding: '2rem',
          borderRadius: '24px',
          border: '1px solid var(--main-border)',
          boxShadow: '0 10px 40px rgba(0,0,0,0.1)'
        }}>
          <h4 style={{ margin: '0 0 1.5rem 0', color: 'var(--table-primary)', fontSize: '1.2rem', fontWeight: 900, display: 'flex', alignItems: 'center', gap: '0.5rem' }}>
            <span>💡</span> Guide Stratégique
          </h4>
          <div style={{ display: 'flex', flexDirection: 'column', gap: '1.25rem' }}>
            {[
              { title: 'Acteurs Clés', color: 'hsla(45, 98%, 51%, 1)', text: 'Haut pouvoir et haut intérêt. Partenaires stratégiques à impliquer étroitement dans toutes les décisions majeures.' },
              { title: 'Garder Satisfait', color: 'hsla(82, 63%, 50%, 1)', text: 'Haut pouvoir mais faible intérêt. Maintenir une satisfaction élevée pour prévenir toute opposition bloquante.' },
              { title: 'Garder Informé', color: 'hsla(25, 95%, 53%, 1)', text: 'Faible pouvoir mais haut intérêt. Les tenir informés régulièrement pour bénéficier de leur soutien et de leur expertise.' },
              { title: 'Effort Minimal', color: 'hsla(217, 91%, 60%, 1)', text: 'Faible pouvoir et faible intérêt. Communication standard et surveillance périodique sans surcharge.' }
            ].map((item, idx) => (
              <div key={idx} style={{
                borderLeft: `4px solid ${item.color}`,
                padding: '0.5rem 0 0.5rem 1.25rem',
                animation: `fadeIn 0.5s ease ${idx * 0.1}s forwards`,
                opacity: 0
              }}>
                <strong style={{ display: 'block', fontSize: '1rem', color: item.color, marginBottom: '4px' }}>{item.title}</strong>
                <p style={{ margin: 0, fontSize: '0.85rem', color: 'var(--text-muted)', lineHeight: '1.5' }}>{item.text}</p>
              </div>
            ))}
          </div>

          <div style={{ marginTop: '2.5rem', padding: '1.25rem', background: 'rgba(59, 130, 246, 0.05)', borderRadius: '16px', border: '1px dashed #3b82f644' }}>
            <div style={{ fontSize: '0.8rem', fontWeight: 700, color: '#3b82f6', marginBottom: '8px' }}>🚀 Automatisation Active</div>
            <div style={{ fontSize: '0.75rem', color: 'var(--text-muted)' }}>Cette matrice se met à jour instantanément lorsque vous modifiez le registre des parties prenantes.</div>
          </div>
        </div>
      </div>
    </div>
  );
};

const PartiesPrenantesSettingsView = () => {
  const [types, setTypes] = useState<string[]>(() => {
    const saved = localStorage.getItem('faab-stakeholder-types');
    if (saved) {
      try {
        const parsed = JSON.parse(saved);
        if (Array.isArray(parsed) && parsed.length > 0) return parsed;
      } catch (e) {}
    }
    return ['Interne', 'Externe'];
  });
  const [newType, setNewType] = useState('');

  const addType = () => {
    if (!newType.trim()) return;
    const updated = [...types, newType.trim()];
    setTypes(updated);
    localStorage.setItem('faab-stakeholder-types', JSON.stringify(updated));
    setNewType('');
  };

  const removeType = (index: number) => {
    const updated = types.filter((_, i) => i !== index);
    setTypes(updated);
    localStorage.setItem('faab-stakeholder-types', JSON.stringify(updated));
  };

  return (
    <div style={{ flex: 1, overflow: 'auto', backgroundColor: 'var(--main-surface)', color: 'var(--main-text)', padding: '2rem', borderRadius: '16px', border: '1px solid var(--main-border)' }}>
      <div style={{ background: 'linear-gradient(135deg, #10b981, #059669)', color: '#fff', padding: '12px 24px', fontSize: '1.2rem', fontWeight: 800, borderRadius: '12px', marginBottom: '2rem' }}>
        ⚙️ Paramétrage du Registre des Parties Prenantes
      </div>

      <div style={{ maxWidth: '480px' }}>
        <h4 style={{ marginBottom: '1rem', color: 'var(--table-primary)', fontWeight: 800 }}>Types de parties prenantes</h4>
        <table style={{ width: '100%', borderCollapse: 'collapse', border: '1px solid var(--main-border)', borderRadius: '12px', overflow: 'hidden' }}>
          <thead>
            <tr style={{ backgroundColor: 'rgba(16, 185, 129, 0.15)', color: 'var(--main-text)' }}>
              <th style={{ padding: '10px 14px', textAlign: 'left', fontWeight: 700 }}>Libellé</th>
              <th style={{ padding: '10px 14px', width: '70px', textAlign: 'center', fontWeight: 700 }}>Action</th>
            </tr>
          </thead>
          <tbody>
            {types.map((t, idx) => (
              <tr key={idx} style={{ borderBottom: '1px solid var(--main-border)' }}>
                <td style={{ padding: '10px 14px', fontWeight: 600 }}>{t}</td>
                <td style={{ textAlign: 'center' }}>
                  <button
                    onClick={() => removeType(idx)}
                    title="Supprimer ce type"
                    style={{ background: 'none', border: 'none', cursor: 'pointer', fontSize: '1.1rem' }}
                  >
                    🗑️
                  </button>
                </td>
              </tr>
            ))}
            <tr>
              <td style={{ padding: '10px 14px' }}>
                <input
                  placeholder="Nouveau type (ex: Fournisseur, Client)..."
                  value={newType}
                  onChange={(e) => setNewType(e.target.value)}
                  onKeyDown={(e) => e.key === 'Enter' && addType()}
                  style={{ width: '100%', padding: '6px 10px', borderRadius: '6px', border: '1px solid var(--main-border)', background: 'rgba(0,0,0,0.02)', color: 'var(--main-text)' }}
                />
              </td>
              <td style={{ textAlign: 'center' }}>
                <button
                  onClick={addType}
                  title="Ajouter"
                  style={{ background: 'var(--table-primary)', color: '#fff', border: 'none', borderRadius: '6px', padding: '6px 12px', cursor: 'pointer', fontWeight: 800 }}
                >
                  ➕
                </button>
              </td>
            </tr>
          </tbody>
        </table>
      </div>
    </div>
  );
};

const RegistreRisquesTableView = ({ fileId, projectId, currentDossierId, initialView, onDossierChange, isVariante1 }: {
  fileId: string;
  projectId: string;
  currentDossierId: string | null;
  initialView?: 'registry' | 'problems' | 'matrix' | 'guide' | 'scales' | 'dashboard' | 'dossiers' | 'settings';
  onDossierChange?: (id: string) => void;
  isVariante1?: boolean;
}) => {
  const [viewMode, setViewMode] = useState(initialView || 'registry');

  useEffect(() => {
    if (initialView) setViewMode(initialView);
  }, [initialView]);

  const dossierSubKey = currentDossierId ? `-${currentDossierId}` : "";
  const storageKey = `registre-risques-${projectId}-${fileId}${dossierSubKey}`;

  const defaultData = [
    { id: '1', phase: 'Phase 1', activite: 'Cadrage', tache: 'Définir objectifs', risque: 'Mauvaise définition des objectifs', impact: 4, probabilite: 3, response: 'A atténuer', action: 'Ateliers de validation', responsable: 'Untel 1', statut: 'En cours' },
    { id: '2', phase: 'Phase 1', activite: 'Planning', tache: 'Créer macro-planning', risque: 'Retard de livraison', impact: 5, probabilite: 2, response: 'A transférer', action: 'Pénalités contrat', responsable: 'Untel 2', statut: 'Fait' },
    { id: '3', phase: 'Phase 2', activite: 'Développement', tache: 'Coder app', risque: 'Bug critique en prod', impact: 5, probabilite: 4, response: 'A atténuer', action: 'Tests rigoureux', responsable: 'Untel 3', statut: 'Non fait' },
  ];

  const [rows, setRows] = useState(() => {
    const saved = localStorage.getItem(storageKey);
    return saved ? JSON.parse(saved) : defaultData;
  });

  useEffect(() => {
    localStorage.setItem(storageKey, JSON.stringify(rows));
  }, [rows, storageKey]);

  useEffect(() => {
    const saved = localStorage.getItem(storageKey);
    if (saved) setRows(JSON.parse(saved));
    else setRows(defaultData);
  }, [storageKey]);

  const handleChange = (id: string, field: string, val: any) => {
    setRows(rows.map((r: any) => r.id === id ? { ...r, [field]: val } : r));
  };

  const addRow = () => {
    const newId = (rows.length + 1).toString();
    setRows([...rows, { id: newId, phase: '', activite: '', tache: '', risque: 'Nouveau risque', impact: 1, probabilite: 1, response: 'Accepter', action: '', responsable: '', statut: 'Non fait' }]);
  };

  const getCriticiteLabel = (impact: number, probabilite: number) => {
    const score = impact * probabilite;
    if (score >= 15) return { label: 'Critique', color: '#ef4444' }; // Rouge
    if (score >= 8) return { label: 'Elevé', color: '#f97316' }; // Orange
    if (score >= 5) return { label: 'Modéré', color: '#eab308' }; // Jaune
    return { label: 'Vert (Faible/Modéré)', color: '#22c55e' }; // Vert
  };

  const autoFillAction = (id: string) => {
    setRows(rows.map((r: any) => {
      if (r.id === id) {
        let response = 'A atténuer';
        let action = 'Mise en place d\'un plan de surveillance et de mitigation.';
        const score = r.impact * r.probabilite;

        if (score >= 15) {
          response = 'A éliminer';
          action = 'Action immédiate requise : Refondre le processus ou supprimer la cause racine.';
        } else if (score < 4) {
          response = 'Accepter';
          action = 'Risque acceptable en l\'état. Surveillance simple.';
        }

        return { ...r, response, action };
      }
      return r;
    }));
  };

  return (
    <div style={{ flex: 1, display: 'flex', flexDirection: 'column', backgroundColor: 'var(--main-bg)', color: 'var(--main-text)', overflow: 'hidden' }}>
      <div style={{ flex: 1, overflow: 'auto', padding: '1.5rem' }}>

        {viewMode === 'guide' && (
          <div style={{ background: 'var(--main-surface)', borderRadius: '12px', padding: '3rem', boxShadow: '0 4px 6px -1px rgba(0,0,0,0.1)', border: '1px solid var(--main-border)' }}>
            <h2 style={{ color: 'var(--table-primary)', marginBottom: '1.5rem', fontWeight: 900 }}>📘 Guide de Gestion des Risques</h2>
            <p style={{ fontSize: '1.1rem', lineHeight: '1.8', color: 'var(--text-muted)' }}>
              Consultez ici les instructions pour gérer vos risques et problèmes au sein de FAAB PROJET.
            </p>
            <ul style={{ lineHeight: '2.5', marginTop: '2rem', fontSize: '1rem' }}>
              <li>📋 <strong>Registre des risques</strong> : identifier, évaluer et suivre les risques projet</li>
              <li>⚠️ <strong>Registre des problèmes</strong> : documenter et résoudre les problèmes survenus</li>
              <li>📊 <strong>Matrice de criticité</strong> : visualiser probabilité × gravité</li>
              <li>⚙️ <strong>Paramétrage</strong> : configurer les échelles et les listes</li>
            </ul>
          </div>
        )}

        {viewMode === 'scales' && (
          <div style={{ background: 'var(--main-surface)', borderRadius: '12px', padding: '2rem', border: '1px solid var(--main-border)' }}>
            <h3 style={{ marginBottom: '1.5rem' }}>📏 Échelles d'évaluation</h3>
            <p>Définissez ici vos échelles d'impact et de probabilité.</p>
            {/* ... simplified ... */}
          </div>
        )}

        {viewMode === 'problems' && (
          <div style={{ background: 'var(--main-surface)', borderRadius: '12px', padding: '1.5rem', boxShadow: '0 4px 6px -1px rgba(0,0,0,0.1)', border: '1px solid var(--main-border)' }}>
            <h3 style={{ margin: '0 0 1.5rem 0', display: 'flex', alignItems: 'center', gap: '10px' }}>⚠️ Registre des problèmes</h3>
            <div style={{ overflowX: 'auto' }}>
              <table style={{ minWidth: '1000px', width: '100%', borderCollapse: 'collapse', border: '1px solid var(--main-border)' }}>
                <thead style={{ backgroundColor: 'var(--table-primary)', color: '#fff' }}>
                  <tr>
                    <th style={{ padding: '10px', borderRight: '1px solid rgba(255,255,255,0.1)' }}>ID</th>
                    <th style={{ padding: '10px', borderRight: '1px solid rgba(255,255,255,0.1)', textAlign: 'left' }}>Phase / Tâche</th>
                    <th style={{ padding: '10px', borderRight: '1px solid rgba(255,255,255,0.1)', textAlign: 'left' }}>Description</th>
                    <th style={{ padding: '10px', borderRight: '1px solid rgba(255,255,255,0.1)' }}>Réponse</th>
                    <th style={{ padding: '10px', borderRight: '1px solid rgba(255,255,255,0.1)', textAlign: 'left' }}>Actions</th>
                    <th style={{ padding: '10px', width: '60px' }}>Action</th>
                  </tr>
                </thead>
                <tbody>
                  {rows.map((r: any, idx: number) => (
                    <tr key={r.id} style={{ borderBottom: '1px solid var(--main-border)', background: idx % 2 === 0 ? 'rgba(0,0,0,0.01)' : 'transparent' }}>
                      <td style={{ padding: '10px', textAlign: 'center', fontWeight: 'bold' }}>{r.id}</td>
                      <td style={{ padding: '8px' }}>
                        <input placeholder="Tâche..." value={r.tache || r.activite} onChange={(e) => handleChange(r.id, 'tache', e.target.value)} style={{ width: '100%', border: 'none', background: 'transparent', color: 'inherit', fontWeight: 600 }} />
                      </td>
                      <td style={{ padding: '8px' }}>
                        <textarea value={r.risque} onChange={(e) => handleChange(r.id, 'risque', e.target.value)} style={{ width: '100%', border: 'none', background: 'rgba(0,0,0,0.02)', padding: '6px', borderRadius: '4px', color: 'inherit', fontSize: '0.8rem', minHeight: '40px' }} />
                      </td>
                      <td style={{ padding: '8px' }}>
                        <select value={r.response} onChange={(e) => handleChange(r.id, 'response', e.target.value)} style={{ width: '100%', padding: '4px', border: '1px solid var(--main-border)', borderRadius: '4px', background: 'transparent', color: 'inherit' }}>
                          <option value="A atténuer">A atténuer</option>
                          <option value="A transférer">A transférer</option>
                          <option value="A éliminer">A éliminer</option>
                        </select>
                      </td>
                      <td style={{ padding: '8px' }}>
                        <textarea value={r.action} onChange={(e) => handleChange(r.id, 'action', e.target.value)} style={{ width: '100%', border: 'none', background: 'rgba(0,0,0,0.02)', padding: '6px', borderRadius: '4px', color: 'inherit', fontSize: '0.8rem', minHeight: '40px' }} />
                      </td>
                      <td style={{ padding: '8px', textAlign: 'center' }}>
                        <button onClick={() => setRows(rows.filter((it: any) => it.id !== r.id))} style={{ background: 'none', border: 'none', cursor: 'pointer' }}>🗑️</button>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
            <button onClick={addRow} style={{ marginTop: '1.5rem', padding: '10px 24px', background: 'var(--table-primary)', color: '#fff', border: 'none', borderRadius: '8px', cursor: 'pointer', fontWeight: 800 }}>+ Ajouter un élément</button>
          </div>
        )}

        {viewMode === 'registry' && (
          <div style={{ background: 'var(--main-surface)', borderRadius: '12px', padding: '1.5rem', boxShadow: '0 4px 6px -1px rgba(0,0,0,0.1)', border: '1px solid var(--main-border)', overflow: 'hidden', display: 'flex', flexDirection: 'column' }}>
            <h3 style={{ margin: '0 0 1.5rem 0', display: 'flex', alignItems: 'center', gap: '10px' }}>
              📋 Registre des risques
            </h3>
            <div style={{ overflowX: 'auto', flex: 1 }}>
              <table style={{ minWidth: '1800px', width: '100%', borderCollapse: 'collapse', border: '1px solid var(--main-border)', fontSize: '0.75rem' }}>
                <thead style={{ position: 'sticky', top: 0, zIndex: 10 }}>
                  <tr style={{ backgroundColor: 'var(--table-primary)', color: '#fff' }}>
                    <th colSpan={5} style={{ padding: '8px', borderRight: '1px solid rgba(255,255,255,0.2)', borderBottom: '1px solid rgba(255,255,255,0.2)' }}>INFORMATIONS GÉNÉRALES</th>
                    <th colSpan={3} style={{ padding: '8px', borderRight: '1px solid rgba(255,255,255,0.2)', borderBottom: '1px solid rgba(255,255,255,0.2)', backgroundColor: '#3b82f6' }}>IDENTIFICATION</th>
                    <th style={{ padding: '8px', borderRight: '1px solid rgba(255,255,255,0.2)', borderBottom: '1px solid rgba(255,255,255,0.2)', backgroundColor: '#6366f1' }}>STATUT</th>
                    <th style={{ padding: '8px', borderRight: '1px solid rgba(255,255,255,0.2)', borderBottom: '1px solid rgba(255,255,255,0.2)', backgroundColor: '#059669' }}>PROBABILITÉ</th>
                    <th style={{ padding: '8px', borderRight: '1px solid rgba(255,255,255,0.2)', borderBottom: '1px solid rgba(255,255,255,0.2)', backgroundColor: '#059669' }}>GRAVITÉ (IMPACT)</th>
                    <th style={{ padding: '8px', borderRight: '1px solid rgba(255,255,255,0.2)', borderBottom: '1px solid rgba(255,255,255,0.2)', backgroundColor: '#dc2626' }}>CRITICITÉ PxG</th>
                    <th style={{ padding: '8px', borderRight: '1px solid rgba(255,255,255,0.2)', borderBottom: '1px solid rgba(255,255,255,0.2)', backgroundColor: '#0891b2' }}>CONSÉQUENCE SI APPARENT</th>
                    <th style={{ padding: '8px', borderRight: '1px solid rgba(255,255,255,0.2)', borderBottom: '1px solid rgba(255,255,255,0.2)', backgroundColor: '#0891b2' }}>TENDANCE</th>
                    <th colSpan={6} style={{ padding: '8px', borderBottom: '1px solid rgba(255,255,255,0.2)', backgroundColor: '#475569' }}>TRAITEMENT & SUIVI</th>
                  </tr>
                  <tr style={{ backgroundColor: 'var(--main-surface)', color: 'var(--main-text)', borderBottom: '2px solid var(--main-border)' }}>
                    {/* Generales */}
                    <th style={{ padding: '10px 6px', borderRight: '1px solid var(--main-border)', minWidth: '40px' }}>ID</th>
                    <th style={{ padding: '10px 6px', borderRight: '1px solid var(--main-border)', minWidth: '120px' }}>Tâche / Lot</th>
                    <th style={{ padding: '10px 6px', borderRight: '1px solid var(--main-border)', minWidth: '120px' }}>Nature du risque</th>
                    <th style={{ padding: '10px 6px', borderRight: '1px solid var(--main-border)', minWidth: '200px' }}>Description</th>
                    <th style={{ padding: '10px 6px', borderRight: '1px solid var(--main-border)', minWidth: '120px' }}>Cause</th>
                    {/* Identification */}
                    <th style={{ padding: '10px 6px', borderRight: '1px solid var(--main-border)', minWidth: '90px' }}>Date</th>
                    <th style={{ padding: '10px 6px', borderRight: '1px solid var(--main-border)', minWidth: '100px' }}>Découvert par</th>
                    <th style={{ padding: '10px 6px', borderRight: '1px solid var(--main-border)', minWidth: '90px' }}>Statut initial</th>
                    {/* Statut, Probabilité, Gravité, Criticité */}
                    <th style={{ padding: '10px 6px', borderRight: '1px solid var(--main-border)', minWidth: '130px', fontSize: '0.65rem', backgroundColor: 'rgba(59, 130, 246, 0.05)' }}>Statut Actuel</th>
                    <th style={{ padding: '10px 6px', borderRight: '1px solid var(--main-border)', minWidth: '100px', fontSize: '0.65rem', backgroundColor: '#dcfce7', color: '#166534' }}>Probabilité</th>
                    <th style={{ padding: '10px 6px', borderRight: '1px solid var(--main-border)', minWidth: '100px', fontSize: '0.65rem', backgroundColor: '#dcfce7', color: '#166534' }}>Impact</th>
                    <th style={{ padding: '10px 6px', borderRight: '1px solid var(--main-border)', minWidth: '100px', fontSize: '0.65rem', backgroundColor: '#dc2626', color: '#fff' }}>Score PxG</th>
                    {/* Conséquence et Tendance */}
                    <th style={{ padding: '10px 6px', borderRight: '1px solid var(--main-border)', minWidth: '90px' }}>Conséquence si apparent</th>
                    <th style={{ padding: '10px 6px', borderRight: '1px solid var(--main-border)', minWidth: '90px' }}>Tendance</th>
                    {/* Traitement */}
                    <th style={{ padding: '10px 6px', borderRight: '1px solid var(--main-border)', minWidth: '90px' }}>Priorité</th>
                    <th style={{ padding: '10px 6px', borderRight: '1px solid var(--main-border)', minWidth: '120px' }}>Déclencheur</th>
                    <th style={{ padding: '10px 6px', borderRight: '1px solid var(--main-border)', minWidth: '150px' }}>Conséquence</th>
                    <th style={{ padding: '10px 6px', borderRight: '1px solid var(--main-border)', minWidth: '200px' }}>Actions Préventives (Eliminer/Réduire)</th>
                    <th style={{ padding: '10px 6px', borderRight: '1px solid var(--main-border)', minWidth: '200px' }}>Actions Correctives (Atténuer/Transférer)</th>
                    <th style={{ padding: '10px 6px', minWidth: '100px' }}>Responsable</th>
                    <th style={{ padding: '10px 6px', minWidth: '40px' }}>🗑️</th>
                  </tr>
                </thead>
                <tbody>
                  {rows.map((r: any, idx: number) => {
                    const prob = Number(r.probabilite) || 1;
                    const grav = Number(r.impact) || 1;
                    const score = prob * grav;

                    const crit = getCriticiteLabel(grav, prob);
                    let critText = '#fff';
                    if (crit.label === 'Modéré') critText = '#854d0e'; // dark yellow
                    else if (crit.label === 'Vert (Faible/Modéré)') critText = '#166534'; // dark green

                    const currStatut = r.statut || 'Nouveau';
                    const getStatusColor = (status: string) => {
                      switch (status) {
                        case 'Nouveau': return { bg: '#e0e7ff', text: '#3730a3' };
                        case 'En cours':
                        case 'En traitement': return { bg: '#dbeafe', text: '#1e40af' };
                        case 'Traité':
                        case 'Clos': return { bg: '#dcfce7', text: '#166534' };
                        default: return { bg: 'rgba(0,0,0,0.05)', text: 'inherit' };
                      }
                    };
                    const stColor = getStatusColor(currStatut);

                    return (
                      <tr key={r.id} style={{ background: idx % 2 === 0 ? 'rgba(0,0,0,0.01)' : 'transparent', borderBottom: '1px solid var(--main-border)' }}>
                        {/* Générales */}
                        <td style={{ padding: '6px', textAlign: 'center', fontWeight: 'bold', borderRight: '1px solid var(--main-border)' }}>{r.id}</td>
                        <td style={{ padding: '6px', borderRight: '1px solid var(--main-border)' }}>
                          <input placeholder="Tâche/Lot..." value={r.tache || r.activite} onChange={(e) => handleChange(r.id, 'tache', e.target.value)} style={{ width: '100%', border: 'none', background: 'transparent', color: 'inherit', fontWeight: 600, fontSize: '0.75rem' }} />
                        </td>
                        <td style={{ padding: '6px', borderRight: '1px solid var(--main-border)' }}>
                          <select value={r.nature || ''} onChange={(e) => handleChange(r.id, 'nature', e.target.value)} style={{ width: '100%', padding: '4px', border: 'none', background: 'transparent', color: 'inherit', fontSize: '0.75rem', cursor: 'pointer' }}>
                            <option value="">Sélectionner...</option>
                            <option value="Politique">Politique</option>
                            <option value="Économique">Économique</option>
                            <option value="Administratif">Administratif</option>
                            <option value="Financier">Financier</option>
                            <option value="Social">Social</option>
                            <option value="Culturel">Culturel</option>
                            <option value="Technologique">Technologique</option>
                            <option value="Technique">Technique</option>
                            <option value="Écologique">Écologique</option>
                            <option value="Sanitaire">Sanitaire</option>
                            <option value="Juridique">Juridique</option>
                          </select>
                        </td>
                        <td style={{ padding: '6px', borderRight: '1px solid var(--main-border)' }}>
                          <textarea value={r.risque} onChange={(e) => handleChange(r.id, 'risque', e.target.value)} rows={2} style={{ width: '100%', border: 'none', background: 'rgba(0,0,0,0.02)', padding: '4px', borderRadius: '4px', color: 'inherit', fontSize: '0.75rem', resize: 'vertical' }} />
                        </td>
                        <td style={{ padding: '6px', borderRight: '1px solid var(--main-border)' }}>
                          <input placeholder="Cause..." value={r.cause || ''} onChange={(e) => handleChange(r.id, 'cause', e.target.value)} style={{ width: '100%', border: 'none', background: 'transparent', color: 'inherit', fontSize: '0.75rem' }} />
                        </td>

                        {/* Identification */}
                        <td style={{ padding: '6px', borderRight: '1px solid var(--main-border)' }}>
                          <input type="date" value={r.dateId || ''} onChange={(e) => handleChange(r.id, 'dateId', e.target.value)} style={{ width: '100%', border: 'none', background: 'transparent', color: 'inherit', fontSize: '0.75rem' }} />
                        </td>
                        <td style={{ padding: '6px', borderRight: '1px solid var(--main-border)' }}>
                          <input placeholder="Nom..." value={r.decouvertPar || ''} onChange={(e) => handleChange(r.id, 'decouvertPar', e.target.value)} style={{ width: '100%', border: 'none', background: 'transparent', color: 'inherit', fontSize: '0.75rem' }} />
                        </td>
                        <td style={{ padding: '6px', borderRight: '1px solid var(--main-border)' }}>
                          <select value={r.statutInit || ''} onChange={(e) => handleChange(r.id, 'statutInit', e.target.value)} style={{ width: '100%', padding: '4px', border: 'none', background: 'transparent', color: 'inherit', fontSize: '0.75rem', cursor: 'pointer' }}>
                            <option value="">Sélectionner...</option>
                            <option value="Élevé">Élevé</option>
                            <option value="Moyen">Moyen</option>
                            <option value="Faible">Faible</option>
                          </select>
                        </td>

                        {/* Dropdowns (Statut, Probabilité, Gravité) */}
                        <td style={{ padding: '6px', borderRight: '1px solid var(--main-border)' }}>
                          <select value={currStatut} onChange={(e) => handleChange(r.id, 'statut', e.target.value)} style={{ width: '100%', padding: '6px', border: '1px solid var(--main-border)', borderRadius: '6px', backgroundColor: stColor.bg, color: stColor.text, fontSize: '0.75rem', fontWeight: 600, cursor: 'pointer' }}>
                            <option value="Nouveau">Nouveau</option>
                            <option value="En traitement">En traitement</option>
                            <option value="Traité">Traité</option>
                            <option value="Clos">Clos</option>
                          </select>
                        </td>

                        <td style={{ padding: '6px', borderRight: '1px solid var(--main-border)' }}>
                          <select value={prob} onChange={(e) => handleChange(r.id, 'probabilite', parseInt(e.target.value))} style={{ width: '100%', padding: '6px', border: '1px solid var(--main-border)', borderRadius: '6px', background: 'var(--main-surface)', color: 'inherit', fontSize: '0.75rem', cursor: 'pointer' }}>
                            <option value={1}>1 - Très Faible</option>
                            <option value={2}>2 - Faible</option>
                            <option value={3}>3 - Fort</option>
                            <option value={4}>4 - Très Fort</option>
                            <option value={5}>5 - Extrême</option>
                          </select>
                        </td>

                        <td style={{ padding: '6px', borderRight: '1px solid var(--main-border)' }}>
                          <select value={grav} onChange={(e) => handleChange(r.id, 'impact', parseInt(e.target.value))} style={{ width: '100%', padding: '6px', border: '1px solid var(--main-border)', borderRadius: '6px', background: 'var(--main-surface)', color: 'inherit', fontSize: '0.75rem', cursor: 'pointer' }}>
                            <option value={1}>1 - Très Faible</option>
                            <option value={2}>2 - Faible</option>
                            <option value={3}>3 - Fort</option>
                            <option value={4}>4 - Très Fort</option>
                            <option value={5}>5 - Extrême</option>
                          </select>
                        </td>

                        {/* Criticité (Auto) */}
                        <td style={{ padding: '6px', borderRight: '1px solid var(--main-border)', textAlign: 'center' }}>
                          <div style={{ display: 'inline-flex', alignItems: 'center', justifyContent: 'center', backgroundColor: crit.color, color: critText, padding: '4px 8px', borderRadius: '12px', fontWeight: 800, fontSize: '0.75rem', minWidth: '80px', boxShadow: '0 1px 2px rgba(0,0,0,0.1)' }}>
                            {score} - {crit.label}
                          </div>
                        </td>

                        {/* Conséquence et Tendance */}
                        <td style={{ padding: '6px', borderRight: '1px solid var(--main-border)' }}>
                          <select value={r.coutImpact || ''} onChange={(e) => handleChange(r.id, 'coutImpact', e.target.value)} style={{ width: '100%', padding: '4px', border: 'none', background: 'transparent', color: 'inherit', fontSize: '0.75rem', cursor: 'pointer' }}>
                            <option value="">Sélectionner...</option>
                            <option value="Coût">Coût</option>
                            <option value="Délai">Délai</option>
                            <option value="Qualité">Qualité</option>
                          </select>
                        </td>
                        <td style={{ padding: '6px', borderRight: '1px solid var(--main-border)' }}>
                          <select value={r.provisions || ''} onChange={(e) => handleChange(r.id, 'provisions', e.target.value)} style={{ width: '100%', padding: '4px', border: 'none', background: 'transparent', color: 'inherit', fontSize: '0.75rem', cursor: 'pointer' }}>
                            <option value="">Sélectionner...</option>
                            <option value="Hausse">Hausse</option>
                            <option value="Stable">Stable</option>
                            <option value="Baisse">Baisse</option>
                          </select>
                        </td>

                        {/* Traitement */}
                        <td style={{ padding: '6px', borderRight: '1px solid var(--main-border)' }}>
                          <select value={r.priorite || 'Normale'} onChange={(e) => handleChange(r.id, 'priorite', e.target.value)} style={{ width: '100%', padding: '2px', border: '1px solid var(--main-border)', borderRadius: '4px', background: 'transparent', color: 'inherit', fontSize: '0.75rem' }}>
                            <option value="Basse">Basse</option>
                            <option value="Normale">Normale</option>
                            <option value="Haute">Haute</option>
                            <option value="Urgente">Urgente</option>
                          </select>
                        </td>
                        <td style={{ padding: '6px', borderRight: '1px solid var(--main-border)' }}>
                          <input placeholder="Déclencheur..." value={r.declencheur || ''} onChange={(e) => handleChange(r.id, 'declencheur', e.target.value)} style={{ width: '100%', border: 'none', background: 'transparent', color: 'inherit', fontSize: '0.75rem' }} />
                        </td>
                        <td style={{ padding: '6px', borderRight: '1px solid var(--main-border)' }}>
                          <textarea placeholder="Conséquence..." value={r.consequence || ''} onChange={(e) => handleChange(r.id, 'consequence', e.target.value)} rows={2} style={{ width: '100%', border: 'none', background: 'rgba(0,0,0,0.02)', padding: '4px', borderRadius: '4px', color: 'inherit', fontSize: '0.75rem', resize: 'vertical' }} />
                        </td>
                        <td style={{ padding: '6px', borderRight: '1px solid var(--main-border)' }}>
                          <textarea placeholder="Actions pour éviter..." value={r.actionPrev || ''} onChange={(e) => handleChange(r.id, 'actionPrev', e.target.value)} rows={2} style={{ width: '100%', border: 'none', background: 'rgba(5, 150, 105, 0.05)', padding: '4px', borderRadius: '4px', color: 'inherit', fontSize: '0.75rem', resize: 'vertical' }} />
                        </td>
                        <td style={{ padding: '6px', borderRight: '1px solid var(--main-border)' }}>
                          <textarea placeholder="Actions si avéré..." value={r.action || ''} onChange={(e) => handleChange(r.id, 'action', e.target.value)} rows={2} style={{ width: '100%', border: 'none', background: 'rgba(220, 38, 38, 0.05)', padding: '4px', borderRadius: '4px', color: 'inherit', fontSize: '0.75rem', resize: 'vertical' }} />
                        </td>
                        <td style={{ padding: '6px', borderRight: '1px solid var(--main-border)' }}>
                          <input placeholder="Responsable..." value={r.responsable || ''} onChange={(e) => handleChange(r.id, 'responsable', e.target.value)} style={{ width: '100%', border: 'none', background: 'transparent', color: 'inherit', fontSize: '0.75rem' }} />
                        </td>
                        <td style={{ padding: '6px', textAlign: 'center' }}>
                          <button onClick={() => setRows(rows.filter((it: any) => it.id !== r.id))} style={{ background: 'none', border: 'none', cursor: 'pointer' }}>🗑️</button>
                        </td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>
            <button
              onClick={addRow}
              style={{ marginTop: '1.5rem', padding: '10px 24px', background: 'var(--table-primary)', color: '#fff', border: 'none', borderRadius: '8px', cursor: 'pointer', fontWeight: 800, alignSelf: 'flex-start' }}
            >
              + Ajouter un risque
            </button>
          </div>
        )}


        {viewMode === 'matrix' && (
          <div style={{ background: 'var(--main-surface)', borderRadius: '12px', padding: '2rem', border: '1px solid var(--main-border)', textAlign: 'center' }}>
            <h3 style={{ marginBottom: '2rem' }}>📊 Matrice de criticité</h3>
            <div style={{ display: 'grid', gridTemplateColumns: 'repeat(5, 120px)', gridTemplateRows: 'repeat(5, 100px)', gap: '5px', margin: '0 auto', width: 'fit-content' }}>
              {/* Simple Heatmap representation */}
              {[5, 4, 3, 2, 1].map(imp => (
                [1, 2, 3, 4, 5].map(prob => {
                  const score = imp * prob;
                  const color = score >= 15 ? '#ef4444' : score >= 8 ? '#f97316' : score >= 4 ? '#eab308' : '#22c55e';
                  const count = rows.filter((r: any) => r.impact === imp && r.probabilite === prob).length;
                  return (
                    <div key={`${imp}-${prob}`} style={{ background: color, color: '#fff', display: 'flex', flexDirection: 'column', alignItems: 'center', justifyContent: 'center', borderRadius: '8px', position: 'relative' }}>
                      <span style={{ fontSize: '1.5rem', fontWeight: 900 }}>{count > 0 ? count : ''}</span>
                      <span style={{ fontSize: '0.6rem', opacity: 0.7 }}>P{prob} I{imp}</span>
                    </div>
                  );
                })
              ))}
            </div>
          </div>
        )}

        {viewMode === 'dashboard' && (
          <div style={{ background: 'var(--main-surface)', borderRadius: '12px', padding: '2rem', border: '1px solid var(--main-border)' }}>
            <h3 style={{ marginBottom: '2rem' }}>📈 Dashboard des risques</h3>
            <div style={{ display: 'grid', gridTemplateColumns: 'repeat(3, 1fr)', gap: '1.5rem' }}>
              <div style={{ padding: '2rem', background: 'var(--main-bg)', borderRadius: '12px', textAlign: 'center', border: '1px solid var(--main-border)' }}>
                <div style={{ fontSize: '0.8rem', color: 'var(--text-muted)' }}>TOTAL RISQUES</div>
                <div style={{ fontSize: '2.5rem', fontWeight: 900 }}>{rows.length}</div>
              </div>
              <div style={{ padding: '2rem', background: 'var(--main-bg)', borderRadius: '12px', textAlign: 'center', border: '1px solid var(--main-border)' }}>
                <div style={{ fontSize: '0.8rem', color: '#ef4444' }}>CRITIQUES</div>
                <div style={{ fontSize: '2.5rem', fontWeight: 900, color: '#ef4444' }}>{rows.filter((r: any) => r.impact * r.probabilite >= 15).length}</div>
              </div>
              <div style={{ padding: '2rem', background: 'var(--main-bg)', borderRadius: '12px', textAlign: 'center', border: '1px solid var(--main-border)' }}>
                <div style={{ fontSize: '0.8rem', color: '#10b981' }}>TRAITÉS</div>
                <div style={{ fontSize: '2.5rem', fontWeight: 900, color: '#10b981' }}>{rows.filter((r: any) => r.statut === 'Fait').length}</div>
              </div>
            </div>
          </div>
        )}

        {viewMode === 'dossiers' && (
          <SelectionDossiersView
            key={`risq-${projectId}-${fileId}`}
            projectId={projectId}
            fileId={fileId}
            title={`Dossiers des Risques ${isVariante1 ? 'Variante 1' : 'Variante 2'}`}
            onSelectDossier={(id, autoNav) => {
              onDossierChange?.(id);
              if (autoNav) setViewMode('registry');
            }}
            onOpenDossier={(id) => {
              onDossierChange?.(id);
              setViewMode('registry');
            }}
            currentDossierId={currentDossierId}
          />
        )}

      </div>
    </div>
  );
};

const RegistreRisquesDashboardView = ({ fileId, projectId }: { fileId: string; projectId: string }) => {
  const storageKey = `registre-risques-${projectId}-${fileId}`;
  const [rows] = useState(() => {
    const saved = localStorage.getItem(storageKey);
    return saved ? JSON.parse(saved) : [];
  });

  const statsCriticite = { 'Non significatif': 0, 'Modéré': 0, 'Elevé': 0, 'Critique': 0 };
  const statsReponse = { 'Accepter': 0, 'A atténuer': 0, 'A éliminer': 0, 'A transférer': 0 };
  const statsAvancement = { 'Fait': 0, 'En cours': 0, 'Non fait': 0 };

  rows.forEach((r: any) => {
    const score = r.impact * r.probabilite;
    let crit = 'Non significatif';
    if (score >= 15) crit = 'Critique';
    else if (score >= 8) crit = 'Elevé';
    else if (score >= 4) crit = 'Modéré';

    statsCriticite[crit as keyof typeof statsCriticite]++;
    if (statsReponse[r.response as keyof typeof statsReponse] !== undefined) statsReponse[r.response as keyof typeof statsReponse]++;
    if (statsAvancement[r.statut as keyof typeof statsAvancement] !== undefined) statsAvancement[r.statut as keyof typeof statsAvancement]++;
  });

  const totalRisques = rows.length;

  return (
    <div style={{ flex: 1, overflow: 'auto', backgroundColor: '#fff', color: '#000', padding: '1.5rem', borderRadius: '8px' }}>
      <div style={{ backgroundColor: '#62b588', color: '#fff', padding: '12px 20px', borderRadius: '4px 4px 0 0', fontWeight: 'bold', fontSize: '1.2rem', marginBottom: '2rem' }}>
        Dashboard des risques
      </div>

      <div style={{ display: 'grid', gridTemplateColumns: '1fr 2fr', gap: '2rem', marginBottom: '2rem' }}>

        {/* Table 1: Statistiques par risque */}
        <div>
          <h3 style={{ fontSize: '1rem', marginBottom: '0.5rem', color: '#333' }}>Statistiques par risque</h3>
          <table style={{ width: '100%', borderCollapse: 'collapse', textAlign: 'center', fontSize: '0.85rem' }}>
            <thead>
              <tr style={{ backgroundColor: '#469b6f', color: '#fff' }}>
                <th style={{ padding: '8px', border: '1px solid #fff' }}>Criticité des risques</th>
                <th style={{ padding: '8px', border: '1px solid #fff' }}>Nombre</th>
              </tr>
            </thead>
            <tbody>
              {Object.entries(statsCriticite).map(([key, val]) => (
                <tr key={key} style={{ borderBottom: '1px solid #eee' }}>
                  <td style={{ padding: '8px', textAlign: 'left' }}>{key}</td>
                  <td style={{ padding: '8px', fontWeight: 'bold' }}>{val}</td>
                </tr>
              ))}
              <tr style={{ backgroundColor: '#f0fdf4', fontWeight: 'bold' }}>
                <td style={{ padding: '8px', textAlign: 'left' }}>Total</td>
                <td style={{ padding: '8px' }}>{totalRisques}</td>
              </tr>
            </tbody>
          </table>
        </div>

        {/* Table 2: Etat d'avancement croisé */}
        <div>
          <h3 style={{ fontSize: '1rem', marginBottom: '0.5rem', color: '#333' }}>Réponse et Avancement</h3>
          <table style={{ width: '100%', borderCollapse: 'collapse', textAlign: 'center', fontSize: '0.85rem' }}>
            <thead>
              <tr style={{ backgroundColor: '#469b6f', color: '#fff' }}>
                <th style={{ padding: '8px', border: '1px solid #fff' }}>Réponse au risque</th>
                <th style={{ padding: '8px', border: '1px solid #fff' }}>Nombre</th>
                <th style={{ padding: '8px', border: '1px solid #fff' }}>Fait</th>
                <th style={{ padding: '8px', border: '1px solid #fff' }}>En cours</th>
                <th style={{ padding: '8px', border: '1px solid #fff' }}>Non fait</th>
              </tr>
            </thead>
            <tbody>
              {Object.entries(statsReponse).map(([key, val]) => {
                const thatRow = rows.filter((r: any) => r.response === key);
                return (
                  <tr key={key} style={{ borderBottom: '1px solid #eee' }}>
                    <td style={{ padding: '8px', textAlign: 'left' }}>{key}</td>
                    <td style={{ padding: '8px', fontWeight: 'bold' }}>{val}</td>
                    <td style={{ padding: '8px' }}>{thatRow.filter((r: any) => r.statut === 'Fait').length}</td>
                    <td style={{ padding: '8px' }}>{thatRow.filter((r: any) => r.statut === 'En cours').length}</td>
                    <td style={{ padding: '8px' }}>{thatRow.filter((r: any) => r.statut === 'Non fait').length}</td>
                  </tr>
                );
              })}
              <tr style={{ backgroundColor: '#f0fdf4', fontWeight: 'bold' }}>
                <td style={{ padding: '8px', textAlign: 'left', borderTop: '2px solid #469b6f' }}>Total</td>
                <td style={{ padding: '8px', borderTop: '2px solid #469b6f' }}>{totalRisques}</td>
                <td style={{ padding: '8px', borderTop: '2px solid #469b6f' }}>{statsAvancement['Fait']}</td>
                <td style={{ padding: '8px', borderTop: '2px solid #469b6f' }}>{statsAvancement['En cours']}</td>
                <td style={{ padding: '8px', borderTop: '2px solid #469b6f' }}>{statsAvancement['Non fait']}</td>
              </tr>
            </tbody>
          </table>
        </div>

      </div>

      <div style={{ display: 'flex', gap: '2rem', justifyContent: 'space-between' }}>
        <div style={{ flex: 1, border: '1px solid #eee', padding: '1rem', borderRadius: '8px' }}>
          <h4 style={{ textAlign: 'center', marginBottom: '1rem' }}>Criticité des risques</h4>
          <div style={{ display: 'flex', flexDirection: 'column', gap: '0.5rem' }}>
            {Object.entries(statsCriticite).map(([key, val]) => val > 0 && (
              <div key={key} style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                <span style={{ fontSize: '0.85rem' }}>{key}</span>
                <div style={{ flex: 1, margin: '0 1rem', height: '12px', background: '#f1f5f9', borderRadius: '6px', overflow: 'hidden' }}>
                  <div style={{ height: '100%', width: `${(val / totalRisques) * 100}%`, background: key === 'Critique' ? '#ef4444' : key === 'Elevé' ? '#f97316' : key === 'Modéré' ? '#eab308' : '#22c55e' }} />
                </div>
                <span style={{ fontSize: '0.8rem', fontWeight: 'bold', width: '30px', textAlign: 'right' }}>{Math.round((val / totalRisques) * 100)}%</span>
              </div>
            ))}
          </div>
        </div>

        <div style={{ flex: 1, border: '1px solid #eee', padding: '1rem', borderRadius: '8px' }}>
          <h4 style={{ textAlign: 'center', marginBottom: '1rem' }}>Réponse au risque</h4>
          <div style={{ display: 'flex', flexDirection: 'column', gap: '0.5rem' }}>
            {Object.entries(statsReponse).map(([key, val]) => val > 0 && (
              <div key={key} style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                <span style={{ fontSize: '0.85rem' }}>{key}</span>
                <div style={{ flex: 1, margin: '0 1rem', height: '12px', background: '#f1f5f9', borderRadius: '6px', overflow: 'hidden' }}>
                  <div style={{ height: '100%', width: `${(val / totalRisques) * 100}%`, background: '#3b82f6' }} />
                </div>
                <span style={{ fontSize: '0.8rem', fontWeight: 'bold', width: '30px', textAlign: 'right' }}>{Math.round((val / totalRisques) * 100)}%</span>
              </div>
            ))}
          </div>
        </div>

        <div style={{ flex: 1, border: '1px solid #eee', padding: '1rem', borderRadius: '8px' }}>
          <h4 style={{ textAlign: 'center', marginBottom: '1rem' }}>Etat d'avancement</h4>
          <div style={{ display: 'flex', flexDirection: 'column', gap: '0.5rem' }}>
            {Object.entries(statsAvancement).map(([key, val]) => val > 0 && (
              <div key={key} style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                <span style={{ fontSize: '0.85rem' }}>{key}</span>
                <div style={{ flex: 1, margin: '0 1rem', height: '12px', background: '#f1f5f9', borderRadius: '6px', overflow: 'hidden' }}>
                  <div style={{ height: '100%', width: `${(val / totalRisques) * 100}%`, background: key === 'Fait' ? '#22c55e' : key === 'En cours' ? '#f59e0b' : '#94a3b8' }} />
                </div>
                <span style={{ fontSize: '0.8rem', fontWeight: 'bold', width: '30px', textAlign: 'right' }}>{Math.round((val / totalRisques) * 100)}%</span>
              </div>
            ))}
          </div>
        </div>
      </div>

    </div>
  );
};

const RegistreRisquesProblemesView = ({ fileId, projectId, type }: { fileId: string; projectId: string; type: 'risques' | 'problemes' }) => {
  const isProblemes = type === 'problemes';
  const DATA_VERSION = 'v3'; // increment this to reset localStorage with new default data
  const storageKey = `registre-${type}-${projectId}-${fileId}`;
  const versionKey = `${storageKey}-version`;

  const defaultData: any[] = isProblemes ? [
    { id: '1', desc: 'Le responsable du changement n\'a pas encore été assigné au projet SAP30.', impact: 'Le plan de communication est en attente et le registre des parties prenantes est incomplet.', date: '5/1/22', declarer: 'Chantal', priorite: 'High', responsable: 'IT', statut: 'Clôturé', action: 'Le sponsor a été informé et comprend l\'importance de la demande.', dateRes: '5/5/22', resFinale: 'Bernard Dupuis est confirmé comme Responsable du Changement et de la Communication Projet SAP30/CMRL. Bienvenue dans l\'équipe, Bernard !', commentaire: 'Ce problème doit être réglé en priorité pour éviter tout retard dans le démarrage du projet.' },
    { id: '2', desc: 'L\'infrastructure informatique entre le CMRL et le siège de Villacoublay est à revoir.', impact: 'Deux problèmes sont à régler : l\'augmentation de la puissance des serveurs pour supporter la charge nouvelle et le réseau dont l\'utilisation va quadrupler dans les mois qui viennent.', date: '5/2/22', declarer: 'Daniel', priorite: 'High', responsable: 'Business', statut: 'Nouveau', action: 'Une étude a été effectuée par le service informatique et les besoins ont été identifiés. La partie hardware semble sous contrôle, la commande et la livraison des nouvelles lignes est à surveiller de très près.', dateRes: '', resFinale: '', commentaire: 'Daniel Château s\'occupe de ce problème en priorité.' },
    { id: '3', desc: 'Les spécifiques du CMRL ne sont pas totalement documentés et leur impact réel reste encore à clarifier.', impact: 'Un impact sur le périmètre ou la charge de travail de réalisation n\'est pas à écarter.', date: '5/3/22', declarer: 'Jacques', priorite: 'Critique', responsable: 'Provider', statut: 'En traitement', action: 'Nos experts du centre d\'excellence informatique visiteront le centre dans les prochains jours afin de documenter l\'écart entre l\'existant et le futur.', dateRes: '', resFinale: '', commentaire: 'Jacques Mounier a été assigné pour coordonner ces travaux. Bien que le système standard du siège soit considéré complet, il n\'est pas impossible que certaines spécificités du CMRL manquent encore. Cette étude est une procédure normale à ce stade du projet.' }
  ] : [
    { id: 'R1', nature: 'Humain', desc: 'Actions correctives', gravite: 'Grave', prob: 'Improbable', criticite: 'Modéré', consequence: 'Qualité', tendance: 'Baisse', responsable: 'Jacques Mounier', actionPrev: 'Contact avec notre cabinet d\'audit afin d\'expliquer la situation et les implications d\'un contrôle inopiné. Développement d\'un plan de contingence qui sera validé par le comité de direction.', actionCorr: 'Comité de Direction exceptionnel pendant lequel le plan de contingence sera expliqué et discuté. Impacts connus aujourd\'hui : coûts et délais sur le projet sera retardé avec certitude.' },
    { id: 'R2', nature: 'Technique', desc: 'Les investissements nécessaires au projet ne sont pas approuvés', gravite: 'Grave', prob: 'Très probable', criticite: 'Très critique', consequence: 'Contenu', tendance: 'Hausse', responsable: 'Chantal Dupont', actionPrev: 'Approbation de l\'étude d\'opportunité ainsi que de la charte projet', actionCorr: 'Comité de Direction exceptionnel et communication des conséquences pour la société.' },
    { id: 'R3', nature: 'Humain', desc: 'Les syndicats n\'approuvent pas le projet', gravite: 'Catastrophique', prob: 'Très probable', criticite: 'Très critique', consequence: 'Contenu', tendance: 'Baisse', responsable: 'Responsable du changement', actionPrev: 'Communication détaillée, transparente et ciblée.', actionCorr: 'Négociation avec les syndicats et explications détaillées.' },
    { id: 'R4', nature: 'Intrinsèque', desc: 'Les ressources informatiques ne sont pas disponibles à temps plein', gravite: 'Grave', prob: 'Peu probable', criticite: 'Critique', consequence: 'Délais', tendance: 'Stable', responsable: 'Chantal Dupont', actionPrev: 'Implication de Chantal Château, DI aux toutes les étapes du projet, y inclus approbation de la charte. Disponibilité des ressources assignées vérifiée en permanence par le chef de projet.', actionCorr: 'Chantal Dupont en contact direct avec la DI afin d\'évaluer les conséquences et appliquer des actions correctives.' },
    { id: 'R5', nature: 'Humain', desc: 'Les ressources locales (responsable du changement, responsable de l\'intégration) ne sont pas disponibles à temps plein', gravite: 'Majeur', prob: 'Très probable', criticite: 'Critique', consequence: 'Délais', tendance: 'Stable', responsable: 'Jacques Mounier', actionPrev: 'Cette disponibilité est sous la responsabilité du responsable de l\'intégration et du responsable du changement. Les ressources ont déjà été approuvées par les responsables de départements CMRL.', actionCorr: 'Le problème sera soulevé lors du prochain comité de direction s\'il n\'est pas réglé entretemps.' }
  ];

  const colsRisques = [
    { k: 'id', l: 'ID', w: '40px' }, { k: 'nature', l: 'Nature du risque', w: '120px' }, { k: 'desc', l: 'Description', w: '250px' },
    { k: 'gravite', l: 'Gravité', w: '100px' }, { k: 'prob', l: 'Probabilité', w: '100px' }, { k: 'criticite', l: 'Criticité', w: '100px' },
    { k: 'consequence', l: 'Conséquence/Impact', w: '140px' }, { k: 'tendance', l: 'Tendance', w: '100px' }, { k: 'responsable', l: 'Responsable', w: '120px' },
    { k: 'actionPrev', l: 'Actions préventives', w: '250px' }, { k: 'actionCorr', l: 'Actions correctives', w: '250px' }
  ];

  const colsProbs = [
    { k: 'id', l: 'ID', w: '40px' }, { k: 'desc', l: 'Description du problème', w: '250px' }, { k: 'impact', l: 'Sommaire d\'impact', w: '200px' },
    { k: 'date', l: 'Date', w: '100px' }, { k: 'declarer', l: 'Déclaré par', w: '120px' }, { k: 'priorite', l: 'Priorité', w: '100px' },
    { k: 'responsable', l: 'Responsable', w: '120px' }, { k: 'statut', l: 'Statut', w: '120px' }, { k: 'action', l: 'Actions correctives', w: '250px' },
    { k: 'dateRes', l: 'Date de Résolution', w: '120px' }, { k: 'resFinale', l: 'Résolution finale', w: '200px' }, { k: 'commentaire', l: 'Commentaire', w: '200px' }
  ];

  const columns = isProblemes ? colsProbs : colsRisques;

  const [rows, setRows] = useState(() => {
    // If the stored version doesn't match DATA_VERSION, reset to fresh defaults
    const storedVersion = localStorage.getItem(versionKey);
    if (storedVersion !== DATA_VERSION) {
      localStorage.removeItem(storageKey);
      localStorage.setItem(versionKey, DATA_VERSION);
      return defaultData;
    }
    const saved = localStorage.getItem(storageKey);
    return saved ? JSON.parse(saved) : defaultData;
  });

  useEffect(() => {
    localStorage.setItem(storageKey, JSON.stringify(rows));
  }, [rows, storageKey]);

  const handleChange = (id: string, field: string, val: any) => {
    setRows(rows.map((r: any) => r.id === id ? { ...r, [field]: val } : r));
  };

  const getCriticiteColor = (val: string) => {
    if (val === 'Très critique') return '#fecaca'; // Red background with red text
    if (val === 'Critique') return '#ffedd5'; // Orange background
    if (val === 'Modéré') return '#fef3c7'; // Yellow background
    return 'transparent';
  };

  const getCriticiteTextColor = (val: string) => {
    if (val === 'Très critique') return '#dc2626';
    if (val === 'Critique') return '#ea580c';
    if (val === 'Modéré') return '#d97706';
    return '#000';
  };

  const getPrioriteColor = (val: string) => {
    if (val === 'Critique') return '#fecaca';
    if (val === 'High') return '#ffedd5';
    if (val === 'Medium') return '#fef3c7';
    if (val === 'Low') return '#ecfccb';
    return 'transparent';
  };

  const getPrioriteTextColor = (val: string) => {
    if (val === 'Critique') return '#ef4444';
    if (val === 'High') return '#f97316';
    if (val === 'Medium') return '#d97706';
    if (val === 'Low') return '#65a30d';
    return '#000';
  };

  return (
    <div style={{ flex: 1, overflow: 'auto', backgroundColor: '#fff', color: '#000', padding: '1.5rem', borderRadius: '8px', fontSize: '0.8rem' }}>
      <div style={{ display: 'flex', alignItems: 'center', marginBottom: '1.5rem' }}>
        <div style={{ width: '30px', height: '30px', borderRadius: '50%', background: '#4ada8e', display: 'flex', alignItems: 'center', justifyContent: 'center', marginRight: '0.5rem' }}>
          <span style={{ color: '#fff', fontSize: '1rem', fontWeight: 'bold' }}>⟡</span>
        </div>
        <div style={{ backgroundColor: '#62c28b', color: '#fff', padding: '10px 20px', borderRadius: '4px', fontWeight: 'bold', fontSize: '1rem', width: '300px' }}>
          {isProblemes ? 'Registre des problèmes' : 'Registre des risques'}
        </div>
      </div>

      <div style={{ overflowX: 'auto' }}>
        <table style={{ minWidth: '1500px', borderCollapse: 'collapse', border: `1px solid #469b6f`, fontSize: '0.75rem' }}>
          <thead>
            {isProblemes && (
              <tr>
                <th colSpan={6} style={{ backgroundColor: '#469b6f', color: '#fff', padding: '8px', border: '1px solid #fff' }}>Identification</th>
                <th colSpan={6} style={{ backgroundColor: '#62b588', color: '#fff', padding: '8px', border: '1px solid #fff' }}>Résolution</th>
                <th style={{ backgroundColor: '#84cc16', color: '#fff', padding: '8px' }}></th>
              </tr>
            )}
            <tr style={{ backgroundColor: '#cbe4d5', color: '#166534' }}>
              {columns.map(c => <th key={c.k} style={{ padding: '8px', border: '1px solid #fff', textAlign: 'left', width: c.w }}>{c.l}</th>)}
              <th style={{ padding: '8px', border: '1px solid #fff', textAlign: 'center', width: '50px' }}>Suppr.</th>
            </tr>
          </thead>
          <tbody>
            {rows.map((r: any) => (
              <tr key={r.id} style={{ borderBottom: '1px solid #eee' }}>
                {columns.map(c => {
                  let cellStyle: any = { padding: '8px', borderRight: '1px solid #eee' };
                  if (c.k === 'criticite' && !isProblemes) {
                    cellStyle.backgroundColor = getCriticiteColor(r[c.k] || '');
                    cellStyle.color = getCriticiteTextColor(r[c.k] || '');
                    cellStyle.fontWeight = 'bold';
                  }
                  if (c.k === 'priorite' && isProblemes) {
                    cellStyle.backgroundColor = getPrioriteColor(r[c.k] || '');
                    cellStyle.color = getPrioriteTextColor(r[c.k] || '');
                    cellStyle.fontWeight = 'bold';
                  }
                  if (c.k === 'statut' && isProblemes) {
                    const status = r[c.k] || '';
                    if (status === 'Nouveau') {
                      cellStyle.color = '#22c55e';
                      cellStyle.fontWeight = 'bold';
                    } else if (status === 'En traitement') {
                      cellStyle.color = '#3b82f6';
                      cellStyle.backgroundColor = '#eff6ff';
                    } else if (status === 'Clôturé') {
                      cellStyle.fontWeight = 'bold';
                    }
                  }

                  return (
                    <td key={`${r.id}-${c.k}`} style={cellStyle}>
                      {c.k === 'desc' || c.k === 'impact' || c.k === 'actionPrev' || c.k === 'actionCorr' || c.k === 'action' || c.k === 'resFinale' || c.k === 'commentaire' ? (
                        <textarea value={r[c.k]} onChange={(e) => handleChange(r.id, c.k, e.target.value)} style={{ width: '100%', border: 'none', background: 'transparent', minHeight: '60px', resize: 'vertical' }} />
                      ) : (
                        <input value={r[c.k]} onChange={(e) => handleChange(r.id, c.k, e.target.value)} style={{ width: '100%', border: 'none', background: 'transparent', fontWeight: cellStyle.fontWeight, color: cellStyle.color }} />
                      )}
                    </td>
                  )
                })}
                <td style={{ padding: '8px', textAlign: 'center' }}>
                  <button className="btn-delete-row" title="Supprimer la ligne" onClick={() => setRows(rows.filter((x: any) => x.id !== r.id))}>❌</button>
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
      <button onClick={() => setRows([...rows, { id: isProblemes ? `${rows.length + 1}` : `R${rows.length + 1}` }])} style={{ marginTop: '1rem', padding: '8px 16px', background: '#469b6f', color: '#fff', border: 'none', borderRadius: '4px', cursor: 'pointer' }}>+ Ajouter {isProblemes ? 'un problème' : 'un risque'}</button>
    </div>
  );
};

const RegistreRisques1ParametrageView = () => {
  return (
    <div style={{ flex: 1, overflow: 'auto', backgroundColor: '#fff', color: '#000', padding: '2rem', borderRadius: '8px' }}>
      <div style={{ backgroundColor: '#469b6f', color: '#fff', padding: '15px 20px', borderRadius: '4px', fontSize: '1.2rem', fontWeight: 'bold', marginBottom: '2rem' }}>
        Paramétrage
      </div>
      <div style={{ display: 'flex', gap: '4rem' }}>
        <div style={{ flex: 1 }}>
          <h3 style={{ color: '#4472c4', marginBottom: '1rem', textTransform: 'uppercase', fontSize: '1rem' }}>Registre des problèmes »</h3>

          <table style={{ width: '100%', marginBottom: '1.5rem', borderCollapse: 'collapse' }}>
            <thead><tr style={{ background: '#469b6f', color: '#fff' }}><th style={{ padding: '8px', textAlign: 'left' }}>Statut</th></tr></thead>
            <tbody>
              <tr><td style={{ padding: '8px', border: '1px solid #ddd', color: '#22c55e', fontWeight: 'bold' }}>Nouveau</td></tr>
              <tr><td style={{ padding: '8px', border: '1px solid #ddd', background: '#eff6ff', color: '#3b82f6' }}>En traitement</td></tr>
              <tr><td style={{ padding: '8px', border: '1px solid #ddd', fontWeight: 'bold' }}>Clôturé</td></tr>
            </tbody>
          </table>

          <table style={{ width: '100%', marginBottom: '1.5rem', borderCollapse: 'collapse' }}>
            <thead><tr style={{ background: '#469b6f', color: '#fff' }}><th style={{ padding: '8px', textAlign: 'left' }}>Responsable</th></tr></thead>
            <tbody>
              <tr><td style={{ padding: '8px', border: '1px solid #ddd', background: '#f8fafc' }}>IT</td></tr>
              <tr><td style={{ padding: '8px', border: '1px solid #ddd', background: '#f8fafc' }}>Provider</td></tr>
              <tr><td style={{ padding: '8px', border: '1px solid #ddd', background: '#f8fafc' }}>Business</td></tr>
            </tbody>
          </table>

          <table style={{ width: '100%', marginBottom: '1.5rem', borderCollapse: 'collapse' }}>
            <thead><tr style={{ background: '#469b6f', color: '#fff' }}><th style={{ padding: '8px', textAlign: 'left' }}>Priorité</th></tr></thead>
            <tbody>
              <tr><td style={{ padding: '8px', border: '1px solid #ddd', background: '#fecaca', color: '#ef4444', fontWeight: 'bold' }}>Critique</td></tr>
              <tr><td style={{ padding: '8px', border: '1px solid #ddd', background: '#ffedd5', color: '#f97316', fontWeight: 'bold' }}>High</td></tr>
              <tr><td style={{ padding: '8px', border: '1px solid #ddd', background: '#fef3c7', color: '#d97706', fontWeight: 'bold' }}>Medium</td></tr>
              <tr><td style={{ padding: '8px', border: '1px solid #ddd', background: '#ecfccb', color: '#65a30d', fontWeight: 'bold' }}>Low</td></tr>
            </tbody>
          </table>
        </div>

        <div style={{ flex: 2 }}>
          <h3 style={{ color: '#4472c4', marginBottom: '1rem', textTransform: 'uppercase', fontSize: '1rem' }}>Registre des risques »</h3>

          <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '2rem', marginBottom: '1.5rem' }}>
            <table style={{ width: '100%', borderCollapse: 'collapse', textAlign: 'center' }}>
              <thead><tr style={{ background: '#469b6f', color: '#fff' }}><th style={{ padding: '8px' }}>Gravité</th><th style={{ padding: '8px' }}>Poids</th></tr></thead>
              <tbody>
                <tr><td style={{ padding: '8px', border: '1px solid #ddd', background: '#f8fafc' }}>Catastrophique</td><td style={{ padding: '8px', border: '1px solid #ddd' }}>4</td></tr>
                <tr><td style={{ padding: '8px', border: '1px solid #ddd', background: '#f8fafc' }}>Grave</td><td style={{ padding: '8px', border: '1px solid #ddd' }}>3</td></tr>
                <tr><td style={{ padding: '8px', border: '1px solid #ddd', background: '#f8fafc' }}>Majeur</td><td style={{ padding: '8px', border: '1px solid #ddd' }}>2</td></tr>
                <tr><td style={{ padding: '8px', border: '1px solid #ddd', background: '#f8fafc' }}>Mineur</td><td style={{ padding: '8px', border: '1px solid #ddd' }}>1</td></tr>
              </tbody>
            </table>

            <table style={{ width: '100%', borderCollapse: 'collapse', textAlign: 'center' }}>
              <thead><tr style={{ background: '#469b6f', color: '#fff' }}><th style={{ padding: '8px' }}>Probabilité</th><th style={{ padding: '8px' }}>Poids</th></tr></thead>
              <tbody>
                <tr><td style={{ padding: '8px', border: '1px solid #ddd', background: '#f8fafc' }}>Très probable</td><td style={{ padding: '8px', border: '1px solid #ddd' }}>4</td></tr>
                <tr><td style={{ padding: '8px', border: '1px solid #ddd', background: '#f8fafc' }}>Probable</td><td style={{ padding: '8px', border: '1px solid #ddd' }}>3</td></tr>
                <tr><td style={{ padding: '8px', border: '1px solid #ddd', background: '#f8fafc' }}>Peu probable</td><td style={{ padding: '8px', border: '1px solid #ddd' }}>2</td></tr>
                <tr><td style={{ padding: '8px', border: '1px solid #ddd', background: '#f8fafc' }}>Improbable</td><td style={{ padding: '8px', border: '1px solid #ddd' }}>1</td></tr>
              </tbody>
            </table>
          </div>

          <table style={{ width: '100%', borderCollapse: 'collapse', textAlign: 'center', marginBottom: '1.5rem' }}>
            <thead><tr style={{ background: '#469b6f', color: '#fff' }}><th style={{ padding: '8px' }}>Min</th><th style={{ padding: '8px' }}>&lt; Criticité ≤</th><th style={{ padding: '8px' }}>Max</th></tr></thead>
            <tbody>
              <tr><td style={{ padding: '8px', border: '1px solid #ddd' }}>-</td><td style={{ padding: '8px', border: '1px solid #ddd', background: '#fef3c7', color: '#d97706', fontWeight: 'bold' }}>Modéré</td><td style={{ padding: '8px', border: '1px solid #ddd' }}>2</td></tr>
              <tr><td style={{ padding: '8px', border: '1px solid #ddd' }}>-</td><td style={{ padding: '8px', border: '1px solid #ddd', background: '#ffedd5', color: '#ea580c', fontWeight: 'bold' }}>Critique</td><td style={{ padding: '8px', border: '1px solid #ddd' }}>8</td></tr>
              <tr><td style={{ padding: '8px', border: '1px solid #ddd' }}>-</td><td style={{ padding: '8px', border: '1px solid #ddd', background: '#fecaca', color: '#dc2626', fontWeight: 'bold' }}>Très critique</td><td style={{ padding: '8px', border: '1px solid #ddd' }}>-</td></tr>
            </tbody>
          </table>

          <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '2rem' }}>
            <table style={{ width: '100%', borderCollapse: 'collapse' }}>
              <thead><tr style={{ background: '#469b6f', color: '#fff' }}><th style={{ padding: '8px', textAlign: 'left' }}>Conséquence</th></tr></thead>
              <tbody>
                <tr><td style={{ padding: '8px', border: '1px solid #ddd', background: '#f8fafc' }}>Délais</td></tr>
                <tr><td style={{ padding: '8px', border: '1px solid #ddd', background: '#f8fafc' }}>Qualité</td></tr>
                <tr><td style={{ padding: '8px', border: '1px solid #ddd', background: '#f8fafc' }}>Budget</td></tr>
                <tr><td style={{ padding: '8px', border: '1px solid #ddd', background: '#f8fafc' }}>Contenu</td></tr>
              </tbody>
            </table>

            <table style={{ width: '100%', borderCollapse: 'collapse' }}>
              <thead><tr style={{ background: '#469b6f', color: '#fff' }}><th style={{ padding: '8px', textAlign: 'left' }}>Nature</th></tr></thead>
              <tbody>
                <tr><td style={{ padding: '8px', border: '1px solid #ddd', background: '#f8fafc' }}>Humain</td></tr>
                <tr><td style={{ padding: '8px', border: '1px solid #ddd', background: '#f8fafc' }}>Technique</td></tr>
                <tr><td style={{ padding: '8px', border: '1px solid #ddd', background: '#f8fafc' }}>Juridique</td></tr>
                <tr><td style={{ padding: '8px', border: '1px solid #ddd', background: '#f8fafc' }}>Intrinsèque</td></tr>
              </tbody>
            </table>
          </div>
        </div>
      </div>
    </div>
  );
};

const RegistreRisques1EchellesView = () => {
  return (
    <div style={{ flex: 1, overflow: 'auto', backgroundColor: '#fff', color: '#000', padding: '2rem', borderRadius: '8px' }}>
      <div style={{ backgroundColor: '#469b6f', color: '#fff', padding: '15px 20px', borderRadius: '4px', fontSize: '1.2rem', fontWeight: 'bold', marginBottom: '2rem' }}>
        Échelles
      </div>
      <div style={{ display: 'flex', flexDirection: 'column', gap: '2rem' }}>
        <div>
          <h3 style={{ backgroundColor: '#469b6f', color: '#fff', padding: '8px 15px', borderRadius: '4px 4px 0 0', margin: 0, textAlign: 'center', fontSize: '1rem' }}>Échelle de fréquence / probabilité</h3>
          <table style={{ width: '100%', borderCollapse: 'collapse', border: '1px solid #469b6f', fontSize: '0.85rem' }}>
            <thead style={{ backgroundColor: '#cbe4d5', color: '#166534' }}>
              <tr>
                <th style={{ padding: '10px', borderRight: '1px solid #fff' }}>Classe de vraisemblance</th>
                <th style={{ padding: '10px', borderRight: '1px solid #fff' }}>Intitulé générique</th>
                <th style={{ padding: '10px', borderRight: '1px solid #fff' }}>Données quantitatives</th>
                <th style={{ padding: '10px' }}>Données qualitatives</th>
              </tr>
            </thead>
            <tbody style={{ textAlign: 'center' }}>
              <tr style={{ borderBottom: '1px solid #eee' }}><td style={{ padding: '10px' }}>V1 (F)</td><td>Impossible à improbable</td><td>&lt; 1 fois en 5 ans</td><td>Jamais vu</td></tr>
              <tr style={{ borderBottom: '1px solid #eee' }}><td style={{ padding: '10px' }}>V2 (F)</td><td>Peu probable</td><td>1 fois par mois</td><td>Vu dans d'autres organisations</td></tr>
              <tr style={{ borderBottom: '1px solid #eee' }}><td style={{ padding: '10px' }}>V3 (F)</td><td>Probable</td><td>1 fois par semaine</td><td>Survient dans l'organisation</td></tr>
              <tr><td style={{ padding: '10px' }}>V4 (F)</td><td>Très probable à certain</td><td>1 fois par jour</td><td>Vécu dans mon secteur d'activité</td></tr>
            </tbody>
          </table>
        </div>

        <div>
          <h3 style={{ backgroundColor: '#469b6f', color: '#fff', padding: '8px 15px', borderRadius: '4px 4px 0 0', margin: 0, textAlign: 'center', fontSize: '1rem' }}>Échelle de gravité / impact</h3>
          <table style={{ width: '100%', borderCollapse: 'collapse', border: '1px solid #469b6f', fontSize: '0.85rem' }}>
            <thead style={{ backgroundColor: '#cbe4d5', color: '#166534' }}>
              <tr>
                <th style={{ padding: '10px', borderRight: '1px solid #fff', width: '20%' }}>Niveau</th>
                <th style={{ padding: '10px', textAlign: 'left' }}>Description synthétique de la gravité</th>
              </tr>
            </thead>
            <tbody>
              <tr style={{ borderBottom: '1px solid #eee' }}><td style={{ padding: '15px', textAlign: 'center', fontWeight: 'bold' }}>Niveau 1</td><td style={{ padding: '15px' }}>Gravité mineure : conséquences mineures sans impact/préjudice pour le projet (ex : retard simple)</td></tr>
              <tr style={{ borderBottom: '1px solid #eee' }}><td style={{ padding: '15px', textAlign: 'center', fontWeight: 'bold' }}>Niveau 2</td><td style={{ padding: '15px' }}>Gravité majeure : incident avec impact pour notre projet (report, prolongation anormale des délais, non satisfaction d'une exigence)</td></tr>
              <tr style={{ borderBottom: '1px solid #eee' }}><td style={{ padding: '15px', textAlign: 'center', fontWeight: 'bold' }}>Niveau 3</td><td style={{ padding: '15px' }}>Gravité critique : conséquences graves pour notre projet (risque liée a la santé et sécurité des utilisateurs, non respect d'une exigence légale et règlementaire, non satisfaction d'un objectif du projet)</td></tr>
              <tr><td style={{ padding: '15px', textAlign: 'center', fontWeight: 'bold' }}>Niveau 4</td><td style={{ padding: '15px' }}>Gravité catastrophique : conséquences très graves pour le projet (non satisfaction des objectifs du projet, compromission des règlements, atteinte a l'image de l'organisation, perte financière, impact significatif sur l'organisation)</td></tr>
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
};

const RegistreRisques1MatriceView = () => {
  return (
    <div style={{ flex: 1, overflow: 'auto', backgroundColor: '#fff', color: '#000', padding: '2rem', borderRadius: '8px' }}>
      <div style={{ backgroundColor: '#62b588', color: '#fff', padding: '15px 20px', borderRadius: '4px', fontSize: '1.2rem', fontWeight: 'bold', marginBottom: '2rem', display: 'flex', gap: '1rem', alignItems: 'center' }}>
        <img src="https://via.placeholder.com/30?text=B" alt="logo" style={{ borderRadius: '50%' }} />
        Matrice de criticité des risques
      </div>

      <div style={{ display: 'flex', gap: '2rem' }}>
        <div style={{ flex: 1, fontSize: '0.8rem', lineHeight: '1.6' }}>
          <div style={{ backgroundColor: '#eff6ff', color: '#3b82f6', padding: '8px 12px', fontWeight: 'bold', marginBottom: '1rem', borderRadius: '4px' }}>LÉGENDE »</div>
          <table style={{ borderCollapse: 'collapse', width: '100%' }}>
            <tbody>
              <tr><td style={{ fontWeight: 'bold', paddingRight: '1rem', width: '30px' }}>R1</td><td>Actions correctives</td></tr>
              <tr><td style={{ fontWeight: 'bold' }}>R2</td><td>Les investissements nécessaires au projet ne sont pas approuvés</td></tr>
              <tr><td style={{ fontWeight: 'bold' }}>R3</td><td>Les syndicats n'approuvent pas le projet</td></tr>
              <tr><td style={{ fontWeight: 'bold' }}>R4</td><td>Les ressources informatiques ne sont pas disponibles à temps plein</td></tr>
              <tr><td style={{ fontWeight: 'bold' }}>R5</td><td>Les ressources locales ne sont pas disponibles</td></tr>
              <tr><td style={{ fontWeight: 'bold' }}>R6</td><td>Le sponsor n'est disponible et ne supporte pas le projet</td></tr>
              <tr><td style={{ fontWeight: 'bold' }}>R7</td><td>Le nouveau système n'est pas disponible avant la clôture</td></tr>
            </tbody>
          </table>
        </div>

        <div style={{ flex: 2, display: 'flex', flexDirection: 'column', alignItems: 'center' }}>
          <div style={{ display: 'flex', width: '100%', maxWidth: '600px', marginBottom: '4px' }}>
            <div style={{ width: '100px' }}></div>
            <div style={{ flex: 1, textAlign: 'center', background: '#65a30d', color: '#fff', margin: '0 4px', borderRadius: '4px', padding: '8px' }}>Improbable</div>
            <div style={{ flex: 1, textAlign: 'center', background: '#65a30d', color: '#fff', margin: '0 4px', borderRadius: '4px', padding: '8px' }}>Peu probable</div>
            <div style={{ flex: 1, textAlign: 'center', background: '#65a30d', color: '#fff', margin: '0 4px', borderRadius: '4px', padding: '8px' }}>Probable</div>
            <div style={{ flex: 1, textAlign: 'center', background: '#65a30d', color: '#fff', margin: '0 4px', borderRadius: '4px', padding: '8px' }}>Très probable</div>
          </div>

          <div style={{ display: 'flex', width: '100%', maxWidth: '600px', height: '100px', marginBottom: '8px' }}>
            <div style={{ width: '100px', display: 'flex', alignItems: 'center', justifyContent: 'center', background: '#65a30d', color: '#fff', borderRadius: '4px', margin: '4px 4px 4px 0', writingMode: 'vertical-rl', transform: 'rotate(180deg)' }}>Catastrophique</div>
            <div style={{ flex: 1, background: '#fef08a', margin: '4px', borderRadius: '8px', display: 'flex', alignItems: 'center', justifyContent: 'center', fontWeight: 'bold', color: '#854d0e' }}></div>
            <div style={{ flex: 1, background: '#fef08a', margin: '4px', borderRadius: '8px', display: 'flex', alignItems: 'center', justifyContent: 'center', fontWeight: 'bold', color: '#854d0e' }}></div>
            <div style={{ flex: 1, background: '#fed7aa', margin: '4px', borderRadius: '8px', display: 'flex', alignItems: 'center', justifyContent: 'center', fontWeight: 'bold', color: '#c2410c' }}></div>
            <div style={{ flex: 1, background: '#fed7aa', margin: '4px', borderRadius: '8px', display: 'flex', alignItems: 'center', justifyContent: 'center', fontWeight: 'bold', color: '#dc2626' }}>R3</div>
          </div>

          <div style={{ display: 'flex', width: '100%', maxWidth: '600px', height: '100px', marginBottom: '8px' }}>
            <div style={{ width: '100px', display: 'flex', alignItems: 'center', justifyContent: 'center', background: '#65a30d', color: '#fff', borderRadius: '4px', margin: '4px 4px 4px 0', writingMode: 'vertical-rl', transform: 'rotate(180deg)' }}>Grave</div>
            <div style={{ flex: 1, background: '#fde047', margin: '4px', borderRadius: '8px', display: 'flex', alignItems: 'center', justifyContent: 'center', fontWeight: 'bold', color: '#854d0e' }}>R1 R6</div>
            <div style={{ flex: 1, background: '#fde047', margin: '4px', borderRadius: '8px', display: 'flex', alignItems: 'center', justifyContent: 'center', fontWeight: 'bold', color: '#854d0e' }}>R4</div>
            <div style={{ flex: 1, background: '#fed7aa', margin: '4px', borderRadius: '8px', display: 'flex', alignItems: 'center', justifyContent: 'center', fontWeight: 'bold', color: '#c2410c' }}></div>
            <div style={{ flex: 1, background: '#fed7aa', margin: '4px', borderRadius: '8px', display: 'flex', alignItems: 'center', justifyContent: 'center', fontWeight: 'bold', color: '#dc2626' }}>R2</div>
          </div>

          <div style={{ display: 'flex', width: '100%', maxWidth: '600px', height: '100px', marginBottom: '8px' }}>
            <div style={{ width: '100px', display: 'flex', alignItems: 'center', justifyContent: 'center', background: '#65a30d', color: '#fff', borderRadius: '4px', margin: '4px 4px 4px 0', writingMode: 'vertical-rl', transform: 'rotate(180deg)' }}>Majeur</div>
            <div style={{ flex: 1, background: '#fef08a', margin: '4px', borderRadius: '8px', display: 'flex', alignItems: 'center', justifyContent: 'center', fontWeight: 'bold', color: '#854d0e' }}></div>
            <div style={{ flex: 1, background: '#fde047', margin: '4px', borderRadius: '8px', display: 'flex', alignItems: 'center', justifyContent: 'center', fontWeight: 'bold', color: '#854d0e' }}></div>
            <div style={{ flex: 1, background: '#fde047', margin: '4px', borderRadius: '8px', display: 'flex', alignItems: 'center', justifyContent: 'center', fontWeight: 'bold', color: '#854d0e' }}></div>
            <div style={{ flex: 1, background: '#facc15', margin: '4px', borderRadius: '8px', display: 'flex', alignItems: 'center', justifyContent: 'center', fontWeight: 'bold', color: '#a16207' }}>R5</div>
          </div>

          <div style={{ display: 'flex', width: '100%', maxWidth: '600px', height: '100px', marginBottom: '8px' }}>
            <div style={{ width: '100px', display: 'flex', alignItems: 'center', justifyContent: 'center', background: '#65a30d', color: '#fff', borderRadius: '4px', margin: '4px 4px 4px 0', writingMode: 'vertical-rl', transform: 'rotate(180deg)' }}>Mineur</div>
            <div style={{ flex: 1, background: '#fef9c3', margin: '4px', borderRadius: '8px', display: 'flex', alignItems: 'center', justifyContent: 'center', fontWeight: 'bold', color: '#854d0e', fontSize: '0.65rem' }}>R8 R9 R10<br />R11 R12</div>
            <div style={{ flex: 1, background: '#fef08a', margin: '4px', borderRadius: '8px', display: 'flex', alignItems: 'center', justifyContent: 'center', fontWeight: 'bold', color: '#854d0e' }}></div>
            <div style={{ flex: 1, background: '#fde047', margin: '4px', borderRadius: '8px', display: 'flex', alignItems: 'center', justifyContent: 'center', fontWeight: 'bold', color: '#854d0e' }}></div>
            <div style={{ flex: 1, background: '#facc15', margin: '4px', borderRadius: '8px', display: 'flex', alignItems: 'center', justifyContent: 'center', fontWeight: 'bold', color: '#a16207' }}>R7</div>
          </div>

        </div>
      </div>
    </div>
  );
};

const RegistreRisquesSettingsView = () => {
  return (
    <div style={{ flex: 1, overflow: 'auto', backgroundColor: '#fff', color: '#000', padding: '2rem', borderRadius: '8px' }}>
      <div style={{ backgroundColor: '#62b588', color: '#fff', padding: '10px 20px', borderRadius: '4px', fontSize: '1.2rem', fontWeight: 'bold', marginBottom: '2rem' }}>
        Paramétrage - Registre des risques
      </div>

      <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '3rem' }}>

        {/* Table 1 */}
        <div>
          <h3 style={{ fontSize: '1rem', marginBottom: '1rem', color: '#333' }}>Niveau de criticité des risques</h3>
          <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '0.85rem', border: '1px solid #469b6f' }}>
            <thead>
              <tr style={{ backgroundColor: '#469b6f', color: '#fff' }}>
                <th style={{ padding: '8px', border: '1px solid #fff', textAlign: 'left' }}>Criticité</th>
                <th style={{ padding: '8px', border: '1px solid #fff', textAlign: 'left' }}>Description (Score = Impact X Proba)</th>
              </tr>
            </thead>
            <tbody>
              <tr style={{ background: '#f8fafc' }}><td style={{ padding: '8px' }}>Non significatif</td><td style={{ padding: '8px' }}>Score &lt; 4</td></tr>
              <tr style={{ background: '#fff' }}><td style={{ padding: '8px' }}>Modéré</td><td style={{ padding: '8px' }}>Score entre 4 et 7</td></tr>
              <tr style={{ background: '#f8fafc' }}><td style={{ padding: '8px' }}>Elevé</td><td style={{ padding: '8px' }}>Score entre 8 et 14</td></tr>
              <tr style={{ background: '#fff' }}><td style={{ padding: '8px' }}>Critique</td><td style={{ padding: '8px' }}>Score &gt;= 15</td></tr>
            </tbody>
          </table>
        </div>

        {/* Table 2 */}
        <div>
          <h3 style={{ fontSize: '1rem', marginBottom: '1rem', color: '#333' }}>Echelle Impact / Probabilité</h3>
          <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '0.85rem', border: '1px solid #469b6f' }}>
            <thead>
              <tr style={{ backgroundColor: '#469b6f', color: '#fff' }}>
                <th style={{ padding: '8px', border: '1px solid #fff' }}>Note</th>
                <th style={{ padding: '8px', border: '1px solid #fff', textAlign: 'left' }}>Description</th>
              </tr>
            </thead>
            <tbody>
              <tr style={{ background: '#f8fafc' }}><td style={{ padding: '8px', textAlign: 'center', fontWeight: 'bold' }}>1</td><td style={{ padding: '8px' }}>Très Faible</td></tr>
              <tr style={{ background: '#fff' }}><td style={{ padding: '8px', textAlign: 'center', fontWeight: 'bold' }}>2</td><td style={{ padding: '8px' }}>Faible</td></tr>
              <tr style={{ background: '#f8fafc' }}><td style={{ padding: '8px', textAlign: 'center', fontWeight: 'bold' }}>3</td><td style={{ padding: '8px' }}>Moyen</td></tr>
              <tr style={{ background: '#fff' }}><td style={{ padding: '8px', textAlign: 'center', fontWeight: 'bold' }}>4</td><td style={{ padding: '8px' }}>Fort</td></tr>
              <tr style={{ background: '#f8fafc' }}><td style={{ padding: '8px', textAlign: 'center', fontWeight: 'bold' }}>5</td><td style={{ padding: '8px' }}>Très Fort</td></tr>
            </tbody>
          </table>
        </div>

        {/* Table 3 */}
        <div>
          <h3 style={{ fontSize: '1rem', marginBottom: '1rem', color: '#333' }}>Réponse au risque</h3>
          <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '0.85rem', border: '1px solid #469b6f' }}>
            <thead>
              <tr style={{ backgroundColor: '#469b6f', color: '#fff' }}>
                <th style={{ padding: '8px', border: '1px solid #fff', textAlign: 'left' }}>Réponse</th>
              </tr>
            </thead>
            <tbody>
              <tr style={{ background: '#f8fafc' }}><td style={{ padding: '8px' }}>Accepter</td></tr>
              <tr style={{ background: '#fff' }}><td style={{ padding: '8px' }}>A atténuer</td></tr>
              <tr style={{ background: '#f8fafc' }}><td style={{ padding: '8px' }}>A transférer</td></tr>
              <tr style={{ background: '#fff' }}><td style={{ padding: '8px' }}>A éliminer</td></tr>
            </tbody>
          </table>
        </div>

        {/* Table 4 */}
        <div>
          <h3 style={{ fontSize: '1rem', marginBottom: '1rem', color: '#333' }}>Etat d'avancement des actions</h3>
          <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '0.85rem', border: '1px solid #469b6f' }}>
            <thead>
              <tr style={{ backgroundColor: '#469b6f', color: '#fff' }}>
                <th style={{ padding: '8px', border: '1px solid #fff', textAlign: 'left' }}>Statut</th>
              </tr>
            </thead>
            <tbody>
              <tr style={{ background: '#f8fafc' }}><td style={{ padding: '8px' }}>Fait</td></tr>
              <tr style={{ background: '#fff' }}><td style={{ padding: '8px' }}>En cours</td></tr>
              <tr style={{ background: '#f8fafc' }}><td style={{ padding: '8px' }}>Non fait</td></tr>
            </tbody>
          </table>
        </div>

      </div>
    </div>
  );
};

const CartographieDecideursTableView = ({ fileId, projectId, currentDossierId, initialView, onDossierChange }: {
  fileId: string;
  projectId: string;
  currentDossierId: string | null;
  initialView?: 'cartography' | 'graph' | 'dossiers';
  onDossierChange?: (id: string) => void;
}) => {
  const [viewMode, setViewMode] = useState<'cartography' | 'graph' | 'dossiers'>(initialView || 'cartography');

  useEffect(() => {
    if (initialView) setViewMode(initialView);
  }, [initialView]);

  const dossierSubKey = currentDossierId ? `-${currentDossierId}` : "";
  const storageKey = `cartographie-decideurs-${projectId}-${fileId}${dossierSubKey}`;

  const defaultData = [
    { id: '1', acteur: 'Untel 1', role: 'Sponsor projet', influence: 5, soutien: 1, strategie: 'Maintien' },
    { id: '2', acteur: 'Untel 2', role: 'Directeur finance', influence: 5, soutien: 1, strategie: '' },
    { id: '3', acteur: 'Untel 3', role: 'Directeur Financier CMRL', influence: 5, soutien: 1, strategie: '' }
  ];

  const [rows, setRows] = useState(() => {
    const saved = localStorage.getItem(storageKey);
    return saved ? JSON.parse(saved) : defaultData;
  });

  useEffect(() => {
    localStorage.setItem(storageKey, JSON.stringify(rows));
  }, [rows, storageKey]);

  useEffect(() => {
    const saved = localStorage.getItem(storageKey);
    if (saved) setRows(JSON.parse(saved));
    else setRows(defaultData);
  }, [storageKey]);

  const getCategory = (influence: number, soutien: number) => {
    if (influence >= 4) {
      if (soutien >= 4) return 'Leaders du projet';
      if (soutien >= 2) return 'Supporteurs passifs';
      return 'Opposants influents';
    } else {
      if (soutien >= 4) return 'Supporteurs actifs';
      if (soutien >= 2) return 'Neutres';
      return 'Opposants mineurs';
    }
  };

  const handleChange = (id: string, field: string, val: any) => {
    setRows(rows.map((r: any) => r.id === id ? { ...r, [field]: val } : r));
  };

  const addRow = () => {
    const newId = (rows.length + 1).toString();
    setRows([...rows, { id: newId, acteur: `Nouvel acteur`, role: '', influence: 3, soutien: 3, strategie: '' }]);
  };

  return (
    <div style={{ flex: 1, display: 'flex', flexDirection: 'column', backgroundColor: 'var(--main-bg)', color: 'var(--main-text)', overflow: 'hidden' }}>
      <div style={{ flex: 1, overflow: 'auto', padding: '1.5rem' }}>

        {viewMode === 'cartography' && (
          <div style={{ background: 'var(--main-surface)', borderRadius: '12px', padding: '1.5rem', boxShadow: '0 4px 6px -1px rgba(0,0,0,0.1)', border: '1px solid var(--main-border)' }}>
            <h3 style={{ margin: '0 0 1.5rem 0' }}>📋 Cartographie des décideurs</h3>
            <div style={{ overflowX: 'auto' }}>
              <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '0.85rem' }}>
                <thead style={{ backgroundColor: 'var(--table-primary)', color: '#fff' }}>
                  <tr>
                    <th style={{ padding: '12px', width: '60px', borderRight: '1px solid rgba(255,255,255,0.1)' }}>ID</th>
                    <th style={{ padding: '12px', textAlign: 'left', borderRight: '1px solid rgba(255,255,255,0.1)' }}>Acteur</th>
                    <th style={{ padding: '12px', textAlign: 'left', borderRight: '1px solid rgba(255,255,255,0.1)' }}>Rôle/fonction</th>
                    <th style={{ padding: '12px', width: '120px', borderRight: '1px solid rgba(255,255,255,0.1)' }}>Influence (1-5)</th>
                    <th style={{ padding: '12px', width: '120px', borderRight: '1px solid rgba(255,255,255,0.1)' }}>Soutien (1-5)</th>
                    <th style={{ padding: '12px', width: '150px', borderRight: '1px solid rgba(255,255,255,0.1)' }}>Catégorie</th>
                    <th style={{ padding: '12px', borderRight: '1px solid rgba(255,255,255,0.1)' }}>Stratégie</th>
                    <th style={{ padding: '12px', width: '60px' }}>Action</th>
                  </tr>
                </thead>
                <tbody>
                  {rows.map((r: any, idx: number) => (
                    <tr key={r.id} style={{ borderBottom: '1px solid var(--main-border)', background: idx % 2 === 0 ? 'rgba(0,0,0,0.01)' : 'transparent' }}>
                      <td style={{ padding: '10px', textAlign: 'center', backgroundColor: 'var(--main-bg)', fontWeight: 'bold' }}>{r.id}</td>
                      <td style={{ padding: '8px' }}><input value={r.acteur} onChange={(e) => handleChange(r.id, 'acteur', e.target.value)} style={{ width: '100%', border: 'none', background: 'rgba(0,0,0,0.02)', padding: '6px', borderRadius: '4px', color: 'inherit', fontWeight: 600 }} /></td>
                      <td style={{ padding: '8px' }}><input value={r.role} onChange={(e) => handleChange(r.id, 'role', e.target.value)} style={{ width: '100%', border: 'none', background: 'rgba(0,0,0,0.02)', padding: '6px', borderRadius: '4px', color: 'inherit' }} /></td>
                      <td style={{ padding: '8px' }}>
                        <select value={r.influence} onChange={(e) => handleChange(r.id, 'influence', parseInt(e.target.value))} style={{ width: '100%', padding: '6px', borderRadius: '4px', border: '1px solid var(--main-border)', background: 'var(--main-surface)', color: 'inherit' }}>
                          {[1, 2, 3, 4, 5].map(n => <option key={n} value={n}>{n}</option>)}
                        </select>
                      </td>
                      <td style={{ padding: '8px' }}>
                        <select value={r.soutien} onChange={(e) => handleChange(r.id, 'soutien', parseInt(e.target.value))} style={{ width: '100%', padding: '6px', borderRadius: '4px', border: '1px solid var(--main-border)', background: 'var(--main-surface)', color: 'inherit' }}>
                          {[1, 2, 3, 4, 5].map(n => <option key={n} value={n}>{n}</option>)}
                        </select>
                      </td>
                      <td style={{ padding: '10px', textAlign: 'center', fontWeight: 'bold' }}>{getCategory(r.influence, r.soutien)}</td>
                      <td style={{ padding: '8px' }}><input value={r.strategie} onChange={(e) => handleChange(r.id, 'strategie', e.target.value)} style={{ width: '100%', border: 'none', background: 'rgba(0,0,0,0.02)', padding: '6px', borderRadius: '4px', color: 'inherit' }} /></td>
                      <td style={{ padding: '8px', textAlign: 'center' }}>
                        <button onClick={() => setRows(rows.filter((it: any) => it.id !== r.id))} style={{ background: 'none', border: 'none', cursor: 'pointer', fontSize: '1.1rem' }}>🗑️</button>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
            <button
              onClick={addRow}
              style={{ marginTop: '1.5rem', padding: '10px 24px', background: 'var(--table-primary)', color: '#fff', border: 'none', borderRadius: '8px', cursor: 'pointer', fontWeight: 800 }}
            >
              + Ajouter un acteur
            </button>
          </div>
        )}

        {viewMode === 'graph' && (
          <CartographieDecideursGraphView rows={rows} />
        )}

        {viewMode === 'dossiers' && (
          <SelectionDossiersView
            key={`carto-${projectId}-${fileId}`}
            projectId={projectId}
            fileId={fileId}
            title="Dossiers de Cartographie"
            onSelectDossier={(id, autoNav) => {
              onDossierChange?.(id);
              if (autoNav) setViewMode('cartography');
            }}
            onOpenDossier={(id) => {
              onDossierChange?.(id);
              setViewMode('cartography');
            }}
            currentDossierId={currentDossierId}
          />
        )}

      </div>
    </div>
  );
};

const CartographieDecideursGraphView = ({ rows }: { rows: any[] }) => {
  return (
    <div style={{ flex: 1, overflow: 'auto', backgroundColor: 'var(--main-bg)', color: 'var(--main-text)', padding: '2rem', display: 'flex', flexDirection: 'column' }}>
      <div style={{ backgroundColor: 'var(--main-surface)', padding: '2rem', borderRadius: '24px', border: '1px solid var(--main-border)', boxShadow: '0 8px 32px rgba(0,0,0,0.05)', display: 'flex', flexDirection: 'column', flex: 1 }}>
        <div style={{
          background: 'linear-gradient(135deg, var(--table-primary) 0%, #469b6f 100%)',
          color: '#fff',
          padding: '1rem 2rem',
          borderRadius: '16px',
          fontSize: '1.3rem',
          fontWeight: 900,
          marginBottom: '1.5rem',
          boxShadow: '0 10px 20px rgba(70, 155, 111, 0.2)',
          display: 'flex',
          alignItems: 'center',
          gap: '0.75rem'
        }}>
          <span>📊</span> Cartographie Stratégique des Décideurs
        </div>

        <div style={{
          fontSize: '0.95rem',
          lineHeight: '1.6',
          marginBottom: '2.5rem',
          color: 'var(--text-muted)',
          background: 'rgba(59, 130, 246, 0.05)',
          padding: '1.25rem',
          borderRadius: '12px',
          borderLeft: '4px solid #3b82f6'
        }}>
          Analyse croisée du <strong>pouvoir d'influence</strong> et du <strong>degré de soutien</strong>.
          Les acteurs à droite sont vos alliés naturels, tandis que ceux en haut à gauche requièrent une gestion proactive de la relation.
        </div>

        <div style={{ flex: 1, display: 'grid', gridTemplateColumns: 'minmax(500px, 1fr) 300px', gap: '3rem', alignItems: 'flex-start' }}>
          <div style={{ position: 'relative', width: '500px', height: '500px', margin: '0 auto', background: 'rgba(0,0,0,0.02)', borderRadius: '20px', border: '1px solid var(--main-border)', padding: '20px' }}>
            {/* Label Y */}
            <div style={{ position: 'absolute', top: '10px', left: '50%', transform: 'translateX(-50%)', fontWeight: 900, fontSize: '0.9rem', color: 'var(--text-muted)', letterSpacing: '2px' }}>INFLUENCE ++</div>
            {/* Label X */}
            <div style={{ position: 'absolute', bottom: '10px', right: '20px', fontWeight: 900, fontSize: '0.9rem', color: 'var(--text-muted)', letterSpacing: '1px' }}>SOUTIEN ++</div>
            <div style={{ position: 'absolute', bottom: '10px', left: '20px', fontWeight: 900, fontSize: '0.9rem', color: 'var(--text-muted)', letterSpacing: '1px' }}>OPPOSANT --</div>

            {/* Grid lines */}
            <div style={{ position: 'absolute', top: '2rem', bottom: '2rem', left: '50%', width: '2px', background: 'repeating-linear-gradient(to bottom, var(--main-border), var(--main-border) 10px, transparent 10px, transparent 20px)' }} />
            <div style={{ position: 'absolute', left: '2rem', right: '2rem', top: '50%', height: '2px', background: 'repeating-linear-gradient(to right, var(--main-border), var(--main-border) 10px, transparent 10px, transparent 20px)' }} />

            {/* Points */}
            {rows.map((r: any) => {
              const x = ((r.soutien - 1) / 4) * 100; // 1 -> 0%, 5 -> 100%
              const y = (1 - (r.influence - 1) / 4) * 100; // 5 -> 0%, 1 -> 100%

              return (
                <div key={r.id} style={{
                  position: 'absolute',
                  left: `${x}%`,
                  top: `${y}%`,
                  transform: 'translate(-50%, -50%)',
                  width: '32px',
                  height: '32px',
                  backgroundColor: 'var(--table-primary)',
                  borderRadius: '50%',
                  display: 'flex',
                  alignItems: 'center',
                  justifyContent: 'center',
                  color: '#fff',
                  fontSize: '0.85rem',
                  fontWeight: 900,
                  boxShadow: '0 4px 10px rgba(0,0,0,0.3)',
                  border: '3px solid #fff',
                  cursor: 'help',
                  transition: 'all 0.2s'
                }}
                  title={`${r.acteur} (${r.role})`}
                  onMouseEnter={(e) => { e.currentTarget.style.transform = 'translate(-50%, -50%) scale(1.4)'; e.currentTarget.style.zIndex = '10'; }}
                  onMouseLeave={(e) => { e.currentTarget.style.transform = 'translate(-50%, -50%) scale(1)'; e.currentTarget.style.zIndex = '1'; }}
                >
                  {r.id}
                </div>
              );
            })}
          </div>

          <div style={{ background: 'var(--main-bg)', padding: '1.5rem', borderRadius: '20px', border: '1px solid var(--main-border)', boxShadow: '0 4px 15px rgba(0,0,0,0.05)' }}>
            <h4 style={{ margin: '0 0 1rem 0', fontSize: '1rem', fontWeight: 800 }}>Liste des Acteurs</h4>
            <div style={{ display: 'flex', flexDirection: 'column', gap: '0.5rem', maxHeight: '400px', overflowY: 'auto', paddingRight: '8px' }}>
              {rows.map((r: any) => (
                <div key={r.id} style={{ display: 'flex', alignItems: 'center', gap: '10px', padding: '8px 12px', background: 'var(--main-surface)', borderRadius: '10px', border: '1px solid var(--main-border)' }}>
                  <div style={{ background: 'var(--table-primary)', color: '#fff', width: '24px', height: '24px', borderRadius: '6px', display: 'flex', alignItems: 'center', justifyContent: 'center', fontWeight: 900, fontSize: '0.75rem' }}>{r.id}</div>
                  <div style={{ flex: 1, fontSize: '0.85rem' }}>
                    <div style={{ fontWeight: 700 }}>{r.acteur}</div>
                    <div style={{ fontSize: '0.7rem', color: 'var(--text-muted)' }}>{r.role}</div>
                  </div>
                </div>
              ))}
            </div>
          </div>
        </div>
      </div>
    </div>
  );
};


const PromptModal = ({ config }: any) => {
  const [val, setVal] = useState('');
  useEffect(() => { setVal(config.defaultValue || '') }, [config]);

  if (!config.isOpen) return null;
  return (
    <div className="viewer-overlay" style={{ zIndex: 200, alignItems: 'center' }}>
      <div className="viewer-content" style={{ height: 'auto', padding: '2.5rem', maxWidth: '450px' }}>
        <h3 style={{ marginBottom: '1.5rem', fontSize: '1.2rem' }}>{config.title}</h3>
        <input
          autoFocus
          className="table-cell-input"
          style={{ width: '100%', background: 'rgba(255,255,255,0.05)', marginBottom: '2rem', fontSize: '1.1rem' }}
          value={val}
          onChange={e => setVal(e.target.value)}
          onKeyDown={e => { if (e.key === 'Enter') config.onConfirm(val); if (e.key === 'Escape') config.onCancel(); }}
        />
        <div style={{ display: 'flex', gap: '1rem', justifyContent: 'flex-end' }}>
          <button className="btn-action" style={{ background: 'transparent' }} onClick={config.onCancel}>Annuler</button>
          <button className="btn-action macro" onClick={() => config.onConfirm(val)}>Valider</button>
        </div>
      </div>
    </div>
  )
}

// PortefeuilleView removed as per requested tab deletion.


const CommunicationActionPlanView = ({ projectId, fileId, currentDossierId, page = 'option1' }: { projectId: string, fileId: string, currentDossierId: string | null, page?: string }) => {
  const metadata = { projectName: 'SAP30 / CMRL', manager: 'Mme Dupont', version: 'V2.3', lastUpdate: '15/05/2023' };
  const dossierSubKey = currentDossierId ? `-${currentDossierId}` : "";
  const storageKey = `faab-comm-data-${projectId}-${fileId}${dossierSubKey}`;
  const defaultData = [
    { target: 'Utilisateurs finaux', message: 'Lancement du module Finances', channel: 'Email / Newsletter', date: '01/09/2023', resp: 'Changement', status: 'Planifié' },
    { target: 'Sponsors', message: 'Rapport d\'avancement mensuel', channel: 'Réunion CODIR', date: '15/09/2023', resp: 'Chef de projet', status: 'Fait' },
    { target: 'Équipe Projet', message: 'Daily Scrum', channel: 'Teams', date: 'Quotidien', resp: 'Scrum Master', status: 'En cours' },
  ];

  const [data, setData] = useState(() => {
    const saved = localStorage.getItem(storageKey);
    return saved ? JSON.parse(saved) : defaultData;
  });

  useEffect(() => {
    localStorage.setItem(storageKey, JSON.stringify(data));
  }, [data, storageKey]);

  if (page === 'option2') {
    return (
      <div style={{ flex: 1, display: 'flex', flexDirection: 'column', background: 'var(--main-bg)' }}>
        <ExecutionHeader title="24 — Engagement des Sponsors" metadata={metadata} />
        <div style={{ padding: '2.5rem', flex: 1 }}>
          <div style={{ background: 'var(--main-surface)', borderRadius: '24px', border: '1px solid var(--main-border)', padding: '3rem', boxShadow: 'var(--card-shadow)', textAlign: 'center' }}>
            <div style={{ fontSize: '4rem', marginBottom: '1.5rem' }}>🎯</div>
            <h3 style={{ color: 'var(--main-text)', fontSize: '1.8rem', fontWeight: 900, marginBottom: '1rem' }}>Focus : Engagement du Sponsor</h3>
            <p style={{ color: 'var(--text-muted)', fontSize: '1.1rem', maxWidth: '600px', margin: '0 auto 2rem' }}>
              Cette vue permet de piloter spécifiquement les actions de communication destinées aux sponsors du projet (Comités de Pilotage, Validations stratégiques).
            </p>
            <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(250px, 1fr))', gap: '2rem' }}>
              {data.filter((d: any) => d.target.toLowerCase().includes('sponsor')).map((d: any, i: number) => (
                <div key={i} style={{ background: 'var(--main-bg)', padding: '1.5rem', borderRadius: '16px', border: '1px solid var(--main-border)', textAlign: 'left' }}>
                  <div style={{ fontWeight: 900, color: 'var(--table-primary)', marginBottom: '0.5rem' }}>{d.message}</div>
                  <div style={{ fontSize: '0.85rem', color: 'var(--text-muted)' }}>Canal: {d.channel}</div>
                  <div style={{ fontSize: '0.85rem', color: 'var(--text-muted)' }}>Responsable: {d.resp}</div>
                  <div style={{ marginTop: '1rem', display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                    <span style={{ fontSize: '0.8rem', fontWeight: 800 }}>{d.date}</span>
                    <span style={{ padding: '4px 10px', borderRadius: '20px', fontSize: '0.7rem', fontWeight: 900, background: d.status === 'Fait' ? '#dcfce7' : '#fef9c3', color: d.status === 'Fait' ? '#166534' : '#854d0e' }}>{d.status}</span>
                  </div>
                </div>
              ))}
            </div>
          </div>
        </div>
      </div>
    );
  }

  return (
    <div style={{ flex: 1, display: 'flex', flexDirection: 'column', background: 'var(--main-bg)' }}>
      <ExecutionHeader title="24 — Plan d'Action Communication" metadata={metadata} />
      <div style={{ flex: 1, overflow: 'auto', padding: '2.5rem' }}>
        <div style={{ background: 'var(--main-surface)', borderRadius: '24px', border: '1px solid var(--main-border)', overflow: 'hidden', boxShadow: '0 8px 32px rgba(0,0,0,0.05)' }}>
          <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '0.85rem' }}>
            <thead style={{ background: 'var(--table-primary)', color: '#fff' }}>
              <tr>
                <th style={{ padding: '15px', textAlign: 'left' }}>Public Cible</th>
                <th style={{ padding: '15px', textAlign: 'left' }}>Objet / Message</th>
                <th style={{ padding: '15px', textAlign: 'left' }}>Canal</th>
                <th style={{ padding: '15px', textAlign: 'left' }}>Date / Fréquence</th>
                <th style={{ padding: '15px', textAlign: 'left' }}>Responsable</th>
                <th style={{ padding: '15px', textAlign: 'center' }}>Statut</th>
              </tr>
            </thead>
            <tbody>
              {data.map((row: any, i: number) => (
                <tr key={i} style={{ borderBottom: '1px solid var(--main-border)' }}>
                  <td style={{ padding: '12px 15px', fontWeight: 800 }}>
                    <input value={row.target} onChange={(e) => setData(data.map((r: any, idx: number) => idx === i ? { ...r, target: e.target.value } : r))} style={{ width: '100%', border: 'none', background: 'transparent', color: 'inherit', fontWeight: 'inherit' }} />
                  </td>
                  <td style={{ padding: '12px 15px' }}>
                    <input value={row.message} onChange={(e) => setData(data.map((r: any, idx: number) => idx === i ? { ...r, message: e.target.value } : r))} style={{ width: '100%', border: 'none', background: 'transparent', color: 'inherit' }} />
                  </td>
                  <td style={{ padding: '12px 15px' }}>
                    <select value={row.channel} onChange={(e) => setData(data.map((r: any, idx: number) => idx === i ? { ...r, channel: e.target.value } : r))} style={{ padding: '4px', borderRadius: '4px', border: '1px solid var(--main-border)', background: 'var(--main-bg)', color: 'inherit' }}>
                      <option>Email / Newsletter</option>
                      <option>Réunion CODIR</option>
                      <option>Teams</option>
                      <option>Atelier</option>
                      <option>Formation</option>
                    </select>
                  </td>
                  <td style={{ padding: '12px 15px' }}>
                    <input value={row.date} onChange={(e) => setData(data.map((r: any, idx: number) => idx === i ? { ...r, date: e.target.value } : r))} style={{ width: '100%', border: 'none', background: 'transparent', color: 'inherit' }} />
                  </td>
                  <td style={{ padding: '12px 15px' }}>
                    <input value={row.resp} onChange={(e) => setData(data.map((r: any, idx: number) => idx === i ? { ...r, resp: e.target.value } : r))} style={{ width: '100%', border: 'none', background: 'transparent', color: 'inherit' }} />
                  </td>
                  <td style={{ padding: '12px 15px', textAlign: 'center' }}>
                    <span style={{ padding: '4px 12px', borderRadius: '20px', fontSize: '0.7rem', fontWeight: 800, background: row.status === 'Fait' ? '#dcfce7' : '#fef9c3', color: row.status === 'Fait' ? '#166534' : '#854d0e' }}>{row.status}</span>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
          <div style={{ padding: '1.5rem' }}>
            <button onClick={() => setData([...data, { target: 'Nouveau public', message: '', channel: 'Email', date: '', resp: '', status: 'Planifié' }])} style={{ padding: '10px 24px', background: 'var(--table-primary)', color: '#fff', border: 'none', borderRadius: '12px', fontWeight: 800, cursor: 'pointer' }}>+ Ajouter une action</button>
          </div>
        </div>
      </div>
    </div>
  );
};

const WBSPBSView = ({ projectId, fileId, currentDossierId, page = 'option1' }: { projectId: string, fileId: string, currentDossierId: string | null, page?: string }) => {
  const dossierSubKey = currentDossierId ? `-${currentDossierId}` : "";
  const storageKey = `faab-wbs-pbs-data-${projectId}-${fileId}${dossierSubKey}`;
  const isWBS = page === 'option2';
  const defaultData = [
    { lvl: 1, id: '1', sap: isWBS ? 'Macro-tâche 1' : 'Module finances', bold: true },
    { lvl: 2, id: '1.1', sap: isWBS ? 'Tâche 1.1' : 'Comptabilité générale' },
    { lvl: 2, id: '1.2', sap: isWBS ? 'Tâche 1.2' : 'Comptabilité client' },
    { lvl: 2, id: '1.3', sap: isWBS ? 'Tâche 1.3' : 'Comptabilité fournisseur' },
    { lvl: 2, id: '1.4', sap: isWBS ? 'Tâche 1.4' : 'Comptabilité des immobilisations' },
    { lvl: 2, id: '1.5', sap: isWBS ? 'Tâche 1.5' : 'Gestion de trésorerie et comptabilité bancaire' },
    { lvl: 1, id: '2', sap: isWBS ? 'Macro-tâche 2' : 'Module achats', bold: true },
    { lvl: 2, id: '2.1', sap: isWBS ? 'Tâche 2.1' : 'Planification des besoins' },
    { lvl: 1, id: '3', sap: isWBS ? 'Macro-tâche 3' : 'Module ventes', bold: true },
    { lvl: 2, id: '3.1', sap: isWBS ? 'Tâche 3.1' : 'Supports de ventes' },
  ];

  const [data, setData] = useState(() => {
    const saved = localStorage.getItem(storageKey);
    return saved ? JSON.parse(saved) : defaultData;
  });

  useEffect(() => {
    localStorage.setItem(storageKey, JSON.stringify(data));
  }, [data, storageKey]);

  const handleChange = (idx: number, field: string, val: any) => {
    const newData = [...data];
    newData[idx] = { ...newData[idx], [field]: val };
    setData(newData);
  };

  return (
    <div style={{ flex: 1, padding: '2rem', background: 'var(--main-bg)', color: 'var(--main-text)', overflow: 'auto' }}>
      <div style={{ maxWidth: '1200px', margin: '0 auto' }}>
        <div style={{ display: 'flex', alignItems: 'center', marginBottom: '2rem', gap: '1rem', background: 'var(--main-surface)', padding: '1.5rem', borderRadius: '16px', border: '1px solid var(--main-border)', boxShadow: '0 4px 12px rgba(0,0,0,0.05)' }}>
          <div style={{ display: 'flex', alignItems: 'center', gap: '1rem', flex: 1 }}>
            <div style={{ width: '40px', height: '40px', borderRadius: '12px', background: 'var(--table-primary)', display: 'flex', alignItems: 'center', justifyContent: 'center', color: '#fff', fontSize: '1.2rem' }}>⟡</div>
            <h2 style={{ margin: 0, fontWeight: 900, color: 'var(--table-primary)' }}>WBS - PBS SAP</h2>
          </div>
          <div style={{ display: 'flex', gap: '0.8rem' }}>
            <div style={{ background: 'var(--table-primary)', color: '#fff', padding: '6px 16px', borderRadius: '30px', fontWeight: 800, fontSize: '0.8rem' }}>PBS SAP</div>
            <div style={{ padding: '6px 16px', background: '#f97316', color: '#fff', borderRadius: '30px', fontWeight: 800, fontSize: '0.8rem' }}>WBS SAP</div>
          </div>
        </div>

        <div style={{ display: 'grid', gridTemplateColumns: '1.5fr 1fr', gap: '2rem', alignItems: 'flex-start' }}>
          <div style={{ background: 'var(--main-surface)', borderRadius: '24px', border: '1px solid var(--main-border)', overflow: 'hidden', boxShadow: '0 8px 32px rgba(0,0,0,0.05)' }}>
            <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '0.85rem' }}>
              <thead style={{ backgroundColor: 'var(--table-primary)', color: '#fff' }}>
                <tr>
                  <th style={{ padding: '12px', width: '80px' }}>Niveau</th>
                  <th style={{ padding: '12px', width: '100px' }}>ID</th>
                  <th style={{ padding: '12px', textAlign: 'left' }}>Désignation SAP</th>
                  <th style={{ padding: '12px', width: '50px' }}></th>
                </tr>
              </thead>
              <tbody>
                {data.map((row: any, i: number) => (
                  <tr key={i} style={{ borderBottom: '1px solid var(--main-border)', background: row.bold ? 'rgba(var(--table-primary-rgb), 0.05)' : 'transparent', fontWeight: row.bold ? 800 : 500 }}>
                    <td style={{ padding: '10px', textAlign: 'center' }}>
                      <input type="number" value={row.lvl} onChange={(e) => handleChange(i, 'lvl', parseInt(e.target.value))} style={{ width: '40px', border: 'none', background: 'transparent', textAlign: 'center', color: 'inherit', fontWeight: 'inherit' }} />
                    </td>
                    <td style={{ padding: '10px', textAlign: 'center', color: 'var(--table-primary)' }}>
                      <input value={row.id} onChange={(e) => handleChange(i, 'id', e.target.value)} style={{ width: '60px', border: 'none', background: 'transparent', textAlign: 'center', color: 'inherit', fontWeight: 900 }} />
                    </td>
                    <td style={{ padding: '10px', paddingLeft: row.lvl > 1 ? `${row.lvl * 20}px` : '15px' }}>
                      <input value={row.sap} onChange={(e) => handleChange(i, 'sap', e.target.value)} style={{ width: '100%', border: 'none', background: 'transparent', color: 'inherit', fontWeight: 'inherit' }} />
                    </td>
                    <td style={{ padding: '10px', textAlign: 'center' }}>
                      <button onClick={() => setData(data.filter((_: any, idx: number) => idx !== i))} style={{ background: 'none', border: 'none', cursor: 'pointer', opacity: 0.5 }}>🗑️</button>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
            <button
              onClick={() => setData([...data, { lvl: 2, id: '', sap: 'Nouvel élément', bold: false }])}
              style={{ margin: '1.5rem', padding: '10px 24px', background: 'var(--table-primary)', color: '#fff', border: 'none', borderRadius: '12px', fontWeight: 800, cursor: 'pointer' }}
            >
              + Ajouter un niveau
            </button>
          </div>

          <div style={{ background: 'rgba(252, 211, 77, 0.1)', padding: '2rem', borderRadius: '24px', border: '2px solid rgba(252, 211, 77, 0.3)', color: 'var(--main-text)' }}>
            <div style={{ display: 'flex', alignItems: 'center', gap: '10px', marginBottom: '1.5rem' }}>
              <span style={{ fontSize: '1.5rem' }}>💡</span>
              <h3 style={{ margin: 0, fontWeight: 900, color: '#b45309' }}>Comprendre le PBS</h3>
            </div>
            <p style={{ fontWeight: 700, lineHeight: 1.6, color: '#92400e' }}>
              Le <strong>PBS (Product Breakdown Structure)</strong> est un procédé de découpage des livrables du projet en sous-produits.
            </p>
            <p style={{ lineHeight: 1.6, opacity: 0.8 }}>
              Dans notre exemple, le progiciel SAP est décomposé en modules fonctionnels :
            </p>
            <ul style={{ lineHeight: 2, fontWeight: 600 }}>
              <li>🏦 Module finances</li>
              <li>📦 Module achats</li>
              <li>🤝 Module ventes</li>
              <li>🚛 Module logistique</li>
            </ul>
          </div>
        </div>
      </div>
    </div>
  );
};

// ─── CompetencesMatriceView ──────────────────────────────────────
// Matrice ressource-compétence : 2 tableaux (Niveau réel + Niveau attendu)
const CompetencesMatriceView = ({ projectId, fileId, currentDossierId, page = 'option1' }: { projectId: string, fileId: string, currentDossierId: string | null, page?: string }) => {
  const dossierSubKey = currentDossierId ? `-${currentDossierId}` : "";
  const storageKey = `faab-comp-matrix-${projectId}-${fileId}${dossierSubKey}`;
  const skillsKey = `faab-comp-skills-${projectId}-${fileId}${dossierSubKey}`;

  if (page === 'option2') return <CompetencesComparaisonView />;
  if (page === 'option3') return <CompetencesDashboardView />;
  if (page === 'option4') {
    return (
      <div style={{ flex: 1, display: 'flex', flexDirection: 'column', background: 'var(--main-bg)', color: 'var(--main-text)', padding: '2rem' }}>
        <div style={{ background: 'var(--main-surface)', padding: '3rem', borderRadius: '32px', border: '1px solid var(--main-border)', boxShadow: 'var(--card-shadow)', textAlign: 'center' }}>
          <div style={{ fontSize: '3rem', marginBottom: '1.5rem' }}>⚙️</div>
          <h3 style={{ fontSize: '1.5rem', fontWeight: 900, marginBottom: '1rem' }}>Paramètres de la Matrice</h3>
          <p style={{ color: 'var(--text-muted)' }}>Configuration des seuils de compétence et des domaines d'évaluation.</p>
        </div>
      </div>
    );
  }

  const defaultSkills = ['Modélisation processus', 'Analyse données', 'Rédaction fonctionnelle', 'Outils métier', 'Scénarios exception', 'Suivi anomalies', 'Reporting', 'Admin CMRL'];
  const defaultResources = [
    { id: 1, name: 'Chantal Dupont', fonc: 'MOA', reel: [2, 1, 1, 3, 1, 4, 1, 1], attendu: [5, 4, 4, 3, 1, 4, 3, 1] },
    { id: 2, name: 'Olivia Robespierre', fonc: 'Finance', reel: [1, 2, 3, 2, 2, 3, 2, 3], attendu: [3, 4, 4, 3, 4, 4, 3, 4] },
    { id: 3, name: 'Jacques Meunier', fonc: 'Logistique', reel: [2, 1, 2, 1, 1, 2, 1, 2], attendu: [3, 3, 3, 3, 3, 3, 3, 3] },
  ];

  const [skills, setSkills] = useState(() => {
    const saved = localStorage.getItem(skillsKey);
    return saved ? JSON.parse(saved) : defaultSkills;
  });

  const [resources, setResources] = useState(() => {
    const saved = localStorage.getItem(storageKey);
    return saved ? JSON.parse(saved) : defaultResources;
  });

  useEffect(() => {
    localStorage.setItem(skillsKey, JSON.stringify(skills));
  }, [skills, skillsKey]);

  useEffect(() => {
    localStorage.setItem(storageKey, JSON.stringify(resources));
  }, [resources, storageKey]);

  const updateSkillValue = (resId: number, type: 'reel' | 'attendu', skillIdx: number, val: number) => {
    setResources(resources.map((r: any) => {
      if (r.id === resId) {
        const newArr = [...r[type]];
        newArr[skillIdx] = val;
        return { ...r, [type]: newArr };
      }
      return r;
    }));
  };

  const Triangle = ({ color = '#3b82f6' }: { color?: string }) => (
    <svg width="10" height="10" viewBox="0 0 10 10" style={{ display: 'inline-block', verticalAlign: 'middle', marginRight: 2 }}>
      <polygon points="0,10 5,0 10,10" fill={color} />
    </svg>
  );

  const renderTable = (title: string, type: 'reel' | 'attendu', triColor: string) => (
    <div style={{ background: 'var(--main-surface)', borderRadius: '20px', border: '1px solid var(--main-border)', overflow: 'hidden', marginBottom: '2rem', boxShadow: '0 4px 20px rgba(0,0,0,0.05)' }}>
      <div style={{ background: 'var(--table-primary)', color: '#fff', padding: '1rem 1.5rem', fontWeight: 900, fontSize: '1rem' }}>{title}</div>
      <div style={{ overflowX: 'auto' }}>
        <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '0.8rem' }}>
          <thead>
            <tr style={{ background: 'rgba(var(--table-primary-rgb), 0.05)', color: 'var(--table-primary)' }}>
              <th style={{ padding: '12px', textAlign: 'left', minWidth: '150px', borderRight: '1px solid var(--main-border)' }}>Ressources</th>
              <th style={{ padding: '12px', textAlign: 'left', minWidth: '120px', borderRight: '1px solid var(--main-border)' }}>Fonction</th>
              {skills.map((s: string, i: number) => (
                <th key={i} style={{ padding: '10px', textAlign: 'center', fontSize: '0.7rem', width: '80px', borderRight: '1px solid var(--main-border)' }}>
                  {s}
                </th>
              ))}
            </tr>
          </thead>
          <tbody>
            {resources.map((r: any, i: number) => (
              <tr key={r.id} style={{ borderBottom: '1px solid var(--main-border)', background: i % 2 === 0 ? 'transparent' : 'rgba(0,0,0,0.01)' }}>
                <td style={{ padding: '10px', fontWeight: 800, borderRight: '1px solid var(--main-border)' }}>{r.name}</td>
                <td style={{ padding: '10px', borderRight: '1px solid var(--main-border)', color: 'var(--text-muted)' }}>{r.fonc}</td>
                {r[type].map((v: number, j: number) => (
                  <td key={j} style={{ padding: '4px', textAlign: 'center', borderRight: '1px solid var(--main-border)', background: v >= 4 ? 'rgba(34, 197, 94, 0.1)' : v >= 3 ? 'rgba(59, 130, 246, 0.1)' : 'transparent' }}>
                    <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'center', gap: '4px' }}>
                      <Triangle color={triColor} />
                      <input
                        type="number"
                        min="0"
                        max="5"
                        value={v}
                        onChange={(e) => updateSkillValue(r.id, type, j, parseInt(e.target.value) || 0)}
                        style={{ width: '25px', border: 'none', background: 'transparent', textAlign: 'center', fontWeight: 900, color: 'inherit' }}
                      />
                    </div>
                  </td>
                ))}
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );

  return (
    <div style={{ flex: 1, padding: '2rem', background: 'var(--main-bg)', color: 'var(--main-text)', overflow: 'auto' }}>
      <div style={{ maxWidth: '1400px', margin: '0 auto' }}>
        <div style={{ display: 'flex', alignItems: 'center', gap: '1rem', marginBottom: '2.5rem' }}>
          <div style={{ width: '40px', height: '40px', borderRadius: '12px', background: 'var(--table-primary)', display: 'flex', alignItems: 'center', justifyContent: 'center', color: '#fff', fontSize: '1.5rem' }}>🎯</div>
          <h2 style={{ margin: 0, fontWeight: 900, color: 'var(--table-primary)', fontSize: '1.8rem' }}>Matrice Ressources-Compétences</h2>
        </div>

        {renderTable("Niveaux de Compétences RÉELS", "reel", "#3b82f6")}
        {renderTable("Niveaux de Compétences ATTENDUS (Cibles)", "attendu", "#9333ea")}

        <div style={{ display: 'flex', gap: '1rem' }}>
          <button onClick={() => setResources([...resources, { id: Date.now(), name: 'Nouvelle ressource', fonc: '', reel: Array(skills.length).fill(0), attendu: Array(skills.length).fill(0) }])} style={{ padding: '12px 24px', background: 'var(--table-primary)', color: '#fff', border: 'none', borderRadius: '12px', fontWeight: 800, cursor: 'pointer' }}>+ Ajouter une ressource</button>
          <button onClick={() => setSkills([...skills, 'Nouvelle compétence'])} style={{ padding: '12px 24px', background: 'var(--main-surface)', color: 'var(--table-primary)', border: '1px solid var(--table-primary)', borderRadius: '12px', fontWeight: 800, cursor: 'pointer' }}>+ Ajouter une compétence</button>
        </div>
      </div>
    </div>
  );
};

// ─── CompetencesComparaisonView ────────────────────────────────────
// Comparaison par ressource (sélectionnable)
const CompetencesComparaisonView = () => {
  const resources = [
    { name: 'Chantal Dupont', fonc: 'MOA', actuel: [2, 1, 1, 3, 1, 4, 1, 1], cible: [4, 3, 3, 4, 3, 4, 3, 3] },
    { name: 'Olivia Robespierre', fonc: 'Finance', actuel: [1, 2, 3, 2, 2, 3, 2, 3], cible: [3, 4, 4, 3, 4, 4, 3, 4] },
    { name: 'Jacques Meunier', fonc: 'Logistique', actuel: [2, 1, 2, 1, 1, 2, 1, 2], cible: [3, 3, 3, 3, 3, 3, 3, 3] },
    { name: 'Monique Davant', fonc: 'Métiers SD', actuel: [1, 3, 1, 2, 3, 1, 2, 1], cible: [3, 4, 3, 4, 4, 3, 4, 3] },
    { name: 'Denise Attali', fonc: 'Achats et Ariba', actuel: [3, 2, 1, 1, 2, 2, 3, 1], cible: [4, 3, 3, 3, 3, 3, 4, 3] },
    { name: 'Jean Albert', fonc: 'Métiers SI', actuel: [2, 1, 3, 2, 1, 3, 1, 2], cible: [3, 3, 4, 4, 3, 4, 3, 4] },
    { name: 'Eric Duteil', fonc: 'Logistique', actuel: [1, 2, 2, 3, 2, 1, 2, 3], cible: [3, 3, 3, 4, 3, 3, 3, 4] },
    { name: 'Daniel Château', fonc: 'Infrastructure Informatique', actuel: [3, 1, 1, 2, 3, 2, 1, 1], cible: [4, 3, 3, 4, 4, 4, 3, 3] },
  ];
  const cols = ['Mod.', 'Analyse', 'Rédac.', 'Outils', 'Scénar.', 'Suivi', 'Reporting', 'CMRL'];
  const colC = (a: number, c: number) => a >= c ? '#16a34a' : '#ef4444';

  return (
    <div style={{ flex: 1, overflow: 'auto', background: '#fff', color: '#000', padding: '1.5rem' }}>
      <div style={{ display: 'flex', alignItems: 'center', marginBottom: '1rem', gap: '0.8rem' }}>
        <div style={{ width: '28px', height: '28px', borderRadius: '50%', background: '#4ada8e', display: 'flex', alignItems: 'center', justifyContent: 'center' }}><span style={{ color: '#fff', fontWeight: 'bold' }}>⟡</span></div>
        <div style={{ background: '#469b6f', color: '#fff', padding: '8px 16px', borderRadius: '4px', fontWeight: 'bold' }}>Comparaison Compétences Actuelles vs Cibles</div>
      </div>
      <div style={{ display: 'flex', gap: '1rem', marginBottom: '1rem', fontSize: '0.8rem' }}>
        <span style={{ display: 'flex', alignItems: 'center', gap: '4px' }}><span style={{ width: '12px', height: '12px', background: '#3b82f6', display: 'inline-block', borderRadius: '2px' }}></span> Niveau actuel</span>
        <span style={{ display: 'flex', alignItems: 'center', gap: '4px' }}><span style={{ width: '12px', height: '12px', background: '#f97316', display: 'inline-block', borderRadius: '2px' }}></span> Niveau cible</span>
        <span style={{ display: 'flex', alignItems: 'center', gap: '4px' }}><span style={{ width: '12px', height: '12px', background: '#16a34a', display: 'inline-block', borderRadius: '2px' }}></span> Objectif atteint</span>
        <span style={{ display: 'flex', alignItems: 'center', gap: '4px' }}><span style={{ width: '12px', height: '12px', background: '#ef4444', display: 'inline-block', borderRadius: '2px' }}></span> Écart à combler</span>
      </div>
      <div style={{ overflowX: 'auto' }}>
        <table style={{ borderCollapse: 'collapse', fontSize: '0.75rem', minWidth: '800px' }}>
          <thead>
            <tr style={{ background: '#469b6f', color: '#fff' }}>
              <th style={{ padding: '8px', border: '1px solid #fff', textAlign: 'left', minWidth: '140px' }}>Ressource</th>
              <th style={{ padding: '8px', border: '1px solid #fff', minWidth: '100px' }}>Fonction</th>
              {cols.map((c, i) => (
                <th key={i} style={{ padding: '6px', border: '1px solid #fff', textAlign: 'center' }}>
                  {c}
                  <div style={{ display: 'flex', justifyContent: 'center', gap: '4px', fontSize: '0.65rem', marginTop: '2px' }}>
                    <span style={{ background: '#3b82f6', padding: '0 4px', borderRadius: '2px' }}>Act.</span>
                    <span style={{ background: '#f97316', padding: '0 4px', borderRadius: '2px' }}>Cib.</span>
                  </div>
                </th>
              ))}
              <th style={{ padding: '8px', border: '1px solid #fff' }}>Écart global</th>
            </tr>
          </thead>
          <tbody>
            {resources.map((r, i) => {
              const totalEcart = r.actuel.reduce((sum, a, j) => sum + (r.cible[j] - a), 0);
              return (
                <tr key={i} style={{ borderBottom: '1px solid #eee', background: i % 2 === 0 ? '#fff' : '#f8fafc' }}>
                  <td style={{ padding: '6px 8px', border: '1px solid #ddd', fontWeight: '500' }}>{r.name}</td>
                  <td style={{ padding: '6px 8px', border: '1px solid #ddd', fontSize: '0.72rem' }}>{r.fonc}</td>
                  {r.actuel.map((a, j) => {
                    const cib = r.cible[j];
                    const ok = a >= cib;
                    return (
                      <td key={j} style={{ padding: '4px', border: '1px solid #ddd', textAlign: 'center' }}>
                        <div style={{ display: 'flex', justifyContent: 'center', gap: '4px' }}>
                          <span style={{ background: '#3b82f6', color: '#fff', borderRadius: '50%', padding: '1px 5px', fontWeight: 'bold', fontSize: '0.7rem' }}>{a}</span>
                          <span style={{ background: '#f97316', color: '#fff', borderRadius: '50%', padding: '1px 5px', fontWeight: 'bold', fontSize: '0.7rem' }}>{cib}</span>
                          <span style={{ color: colC(a, cib), fontWeight: 'bold', fontSize: '0.8rem' }}>{ok ? '✓' : '!'}</span>
                        </div>
                      </td>
                    );
                  })}
                  <td style={{ padding: '6px', border: '1px solid #ddd', textAlign: 'center', fontWeight: 'bold', color: totalEcart === 0 ? '#16a34a' : '#ef4444' }}>
                    {totalEcart > 0 ? `−${totalEcart}` : '✓ OK'}
                  </td>
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>
    </div>
  );
};

// ─── CompetencesDashboardView ──────────────────────────────────────
const CompetencesDashboardView = () => {
  const competences = [
    { name: 'Modélisation processus', actuel: 2.1, cible: 3.5 },
    { name: 'Analyse données', actuel: 1.8, cible: 3.5 },
    { name: 'Rédaction fonctionnelle', actuel: 2.0, cible: 3.0 },
    { name: 'Connaissance outils', actuel: 2.2, cible: 3.5 },
    { name: 'Scénarios exception', actuel: 1.9, cible: 3.0 },
    { name: 'Suivi anomalies', actuel: 2.5, cible: 3.5 },
    { name: 'Conception reporting', actuel: 1.8, cible: 3.0 },
    { name: 'Connaissance CMRL', actuel: 1.7, cible: 3.0 },
  ];

  const resources = [
    { name: 'Chantal Dupont', moy: 2.1, ecart: -1.4 },
    { name: 'Olivia Robespierre', moy: 2.2, ecart: -1.1 },
    { name: 'Jacques Meunier', moy: 1.5, ecart: -1.5 },
    { name: 'Monique Davant', moy: 1.9, ecart: -1.3 },
    { name: 'Denise Attali', moy: 1.9, ecart: -1.1 },
    { name: 'Jean Albert', moy: 1.9, ecart: -1.4 },
    { name: 'Eric Duteil', moy: 2.0, ecart: -1.0 },
    { name: 'Daniel Château', moy: 1.8, ecart: -1.5 },
  ];

  const nTotal = 16;
  const nNecessite = 12;
  const nOK = 4;
  const pct = (v: number, max = 4) => `${(v / max) * 100}%`;

  return (
    <div style={{ flex: 1, overflow: 'auto', background: '#f8fafc', color: '#000', padding: '1.5rem' }}>
      <div style={{ display: 'flex', alignItems: 'center', marginBottom: '1.5rem', gap: '0.8rem' }}>
        <div style={{ width: '28px', height: '28px', borderRadius: '50%', background: '#4ada8e', display: 'flex', alignItems: 'center', justifyContent: 'center' }}><span style={{ color: '#fff', fontWeight: 'bold' }}>⟡</span></div>
        <div style={{ background: '#469b6f', color: '#fff', padding: '8px 16px', borderRadius: '4px', fontWeight: 'bold' }}>Dashboard des Compétences</div>
      </div>

      {/* KPIs */}
      <div style={{ display: 'grid', gridTemplateColumns: 'repeat(3,1fr)', gap: '1rem', marginBottom: '2rem' }}>
        {[
          { label: 'Total ressources', val: nTotal, icon: '👥', color: '#3b82f6' },
          { label: 'Nécessitent formation', val: nNecessite, icon: '⚠️', color: '#f97316' },
          { label: 'Objectifs atteints', val: nOK, icon: '✅', color: '#16a34a' },
        ].map((kpi: { label: string; val: number; icon: string; color: string }, i: number) => (
          <div key={i} style={{ background: '#fff', borderRadius: '12px', padding: '1.2rem', border: '1px solid #e2e8f0', display: 'flex', alignItems: 'center', gap: '1rem', boxShadow: '0 1px 4px rgba(0,0,0,0.05)' }}>
            <span style={{ fontSize: '2rem' }}>{kpi.icon}</span>
            <div>
              <div style={{ fontSize: '2rem', fontWeight: 'bold', color: kpi.color, lineHeight: 1 }}>{kpi.val}</div>
              <div style={{ fontSize: '0.8rem', color: '#64748b' }}>{kpi.label}</div>
            </div>
          </div>
        ))}
      </div>

      <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '2rem' }}>
        {/* Compétences par domaine */}
        <div style={{ background: '#fff', borderRadius: '12px', padding: '1.5rem', border: '1px solid #e2e8f0' }}>
          <h3 style={{ margin: '0 0 1rem', fontSize: '0.95rem', color: '#1e293b' }}>📊 Niveaux par compétence</h3>
          <div style={{ display: 'flex', flexDirection: 'column', gap: '0.8rem' }}>
            {competences.map((c, i) => (
              <div key={i}>
                <div style={{ display: 'flex', justifyContent: 'space-between', marginBottom: '3px', fontSize: '0.78rem' }}>
                  <span>{c.name}</span>
                  <span style={{ fontWeight: 'bold', color: c.actuel >= c.cible ? '#16a34a' : '#f97316' }}>{c.actuel.toFixed(1)} / {c.cible.toFixed(1)}</span>
                </div>
                <div style={{ width: '100%', height: '18px', background: '#f1f5f9', borderRadius: '6px', overflow: 'hidden', position: 'relative' }}>
                  <div style={{ position: 'absolute', left: 0, top: 0, bottom: 0, width: pct(c.cible), background: '#e2e8f0', borderRadius: '6px' }} />
                  <div style={{ position: 'absolute', left: 0, top: 0, bottom: 0, width: pct(c.actuel), background: c.actuel >= c.cible ? '#16a34a' : '#3b82f6', borderRadius: '6px', transition: 'width 0.5s ease' }} />
                </div>
              </div>
            ))}
          </div>
        </div>

        {/* Ressources */}
        <div style={{ background: '#fff', borderRadius: '12px', padding: '1.5rem', border: '1px solid #e2e8f0' }}>
          <h3 style={{ margin: '0 0 1rem', fontSize: '0.95rem', color: '#1e293b' }}>👤 Niveau moyen par ressource</h3>
          <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '0.8rem' }}>
            <thead>
              <tr style={{ background: '#f8fafc' }}>
                <th style={{ padding: '6px', textAlign: 'left', borderBottom: '1px solid #e2e8f0' }}>Ressource</th>
                <th style={{ padding: '6px', textAlign: 'center', borderBottom: '1px solid #e2e8f0' }}>Moy.</th>
                <th style={{ padding: '6px', textAlign: 'center', borderBottom: '1px solid #e2e8f0' }}>Barre</th>
                <th style={{ padding: '6px', textAlign: 'center', borderBottom: '1px solid #e2e8f0' }}>Écart</th>
              </tr>
            </thead>
            <tbody>
              {resources.map((r, i) => (
                <tr key={i} style={{ borderBottom: '1px solid #f1f5f9' }}>
                  <td style={{ padding: '6px', fontWeight: '500' }}>{r.name}</td>
                  <td style={{ padding: '6px', textAlign: 'center', fontWeight: 'bold', color: '#3b82f6' }}>{r.moy}</td>
                  <td style={{ padding: '6px' }}>
                    <div style={{ width: '100%', height: '10px', background: '#f1f5f9', borderRadius: '4px', overflow: 'hidden' }}>
                      <div style={{ height: '100%', width: pct(r.moy), background: '#3b82f6', borderRadius: '4px' }} />
                    </div>
                  </td>
                  <td style={{ padding: '6px', textAlign: 'center', fontWeight: 'bold', color: '#ef4444', fontSize: '0.75rem' }}>{r.ecart}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
};

// ─── CompetencesParametreView ─────────────────────────────────────
const CompetencesParametreView = () => {
  const [competencesList, setCompetencesList] = useState([
    { id: 1, nom: 'Modélisation processus', description: 'Capacité à modéliser les processus métier' },
    { id: 2, nom: 'Analyse données', description: 'Analyse et interprétation des données' },
    { id: 3, nom: 'Rédaction fonctionnelle', description: 'Rédaction de documents fonctionnels' },
    { id: 4, nom: 'Connaissance outils', description: 'Maîtrise des outils du projet' },
    { id: 5, nom: 'Scénarios exception', description: 'Gestion des cas exceptionnels' },
    { id: 6, nom: 'Suivi anomalies', description: 'Suivi et résolution d\'anomalies' },
    { id: 7, nom: 'Conception reporting', description: 'Création de rapports et tableaux de bord' },
    { id: 8, nom: 'Connaissance CMRL', description: 'Connaissance du contexte CMRL' },
  ]);
  const [ressourcesList, setRessourcesList] = useState([
    { id: 1, nom: 'Chantal Dupont', fonction: 'MOA', niveau_cible: 3 },
    { id: 2, nom: 'Olivia Robespierre', fonction: 'Finance', niveau_cible: 3 },
    { id: 3, nom: 'Jacques Meunier', fonction: 'Logistique', niveau_cible: 3 },
    { id: 4, nom: 'Monique Davant', fonction: 'Métiers SD', niveau_cible: 3 },
    { id: 5, nom: 'Denise Attali', fonction: 'Achats et Ariba', niveau_cible: 3 },
    { id: 6, nom: 'Jean Albert', fonction: 'Métiers SI', niveau_cible: 3 },
    { id: 7, nom: 'Eric Duteil', fonction: 'Logistique', niveau_cible: 3 },
    { id: 8, nom: 'Daniel Château', fonction: 'Infrastructure Informatique', niveau_cible: 4 },
  ]);

  const niveauxLabels = ['Débutant', 'Basique', 'Intermédiaire', 'Avancé', 'Expert'];

  return (
    <div style={{ flex: 1, overflow: 'auto', background: '#fff', color: '#000', padding: '2rem' }}>
      <div style={{ display: 'flex', alignItems: 'center', marginBottom: '1.5rem', gap: '0.8rem' }}>
        <div style={{ width: '28px', height: '28px', borderRadius: '50%', background: '#4ada8e', display: 'flex', alignItems: 'center', justifyContent: 'center' }}><span style={{ color: '#fff', fontWeight: 'bold' }}>⟡</span></div>
        <div style={{ background: '#469b6f', color: '#fff', padding: '8px 16px', borderRadius: '4px', fontWeight: 'bold' }}>⚙️ Paramètre — Matrice des compétences</div>
      </div>

      {/* Échelle des niveaux */}
      <div style={{ marginBottom: '2rem' }}>
        <h3 style={{ color: '#469b6f', marginBottom: '0.8rem', fontSize: '0.95rem', display: 'flex', alignItems: 'center', gap: '0.5rem' }}>
          <span>📏</span> Échelle de notation
        </h3>
        <table style={{ borderCollapse: 'collapse', border: '1px solid #469b6f', fontSize: '0.85rem' }}>
          <thead style={{ background: '#469b6f', color: '#fff' }}>
            <tr>
              <th style={{ padding: '8px 16px', border: '1px solid #fff' }}>Note</th>
              <th style={{ padding: '8px 16px', border: '1px solid #fff' }}>Niveau</th>
              <th style={{ padding: '8px 16px', border: '1px solid #fff', textAlign: 'left' }}>Description</th>
            </tr>
          </thead>
          <tbody>
            {niveauxLabels.map((n, i) => (
              <tr key={i} style={{ borderBottom: '1px solid #d1fae5', background: i % 2 === 0 ? '#fff' : '#f0fdf4' }}>
                <td style={{ padding: '8px 16px', textAlign: 'center', fontWeight: 'bold', color: i + 1 >= 4 ? '#16a34a' : i + 1 >= 3 ? '#65a30d' : i + 1 >= 2 ? '#ca8a04' : '#ea580c' }}>{i + 1}</td>
                <td style={{ padding: '8px 16px', fontWeight: '500' }}>{n}</td>
                <td style={{ padding: '8px 16px', color: '#555', fontSize: '0.8rem' }}>
                  {i === 0 ? 'Aucune ou très faible connaissance du domaine' : i === 1 ? 'Connaissance basique, nécessite encadrement' : i === 2 ? "Autonome sur les tâches courantes, maîtrise partielle" : i === 3 ? "Maîtrise complète, peut guider les autres" : "Expert reconnu, peut former et innover"}
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>

      <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '2rem' }}>
        {/* Liste des compétences */}
        <div>
          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '0.8rem' }}>
            <h3 style={{ color: '#469b6f', margin: 0, fontSize: '0.95rem' }}>🎯 Compétences configurées ({competencesList.length})</h3>
            <button onClick={() => setCompetencesList([...competencesList, { id: Date.now(), nom: 'Nouvelle compétence', description: '' }])} style={{ padding: '4px 10px', background: '#469b6f', color: '#fff', border: 'none', borderRadius: '4px', cursor: 'pointer', fontSize: '0.8rem' }}>+ Ajouter</button>
          </div>
          <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '0.8rem', border: '1px solid #d1fae5' }}>
            <thead style={{ background: '#cbe4d5', color: '#166534' }}>
              <tr>
                <th style={{ padding: '8px', textAlign: 'left', border: '1px solid #fff' }}>#</th>
                <th style={{ padding: '8px', textAlign: 'left', border: '1px solid #fff' }}>Compétence</th>
                <th style={{ padding: '8px', textAlign: 'center', border: '1px solid #fff' }}>Action</th>
              </tr>
            </thead>
            <tbody>
              {competencesList.map((c, i) => (
                <tr key={c.id} style={{ borderBottom: '1px solid #f0fdf4' }}>
                  <td style={{ padding: '6px 8px', textAlign: 'center', color: '#469b6f', fontWeight: 'bold' }}>{i + 1}</td>
                  <td style={{ padding: '6px 8px' }}>
                    <input value={c.nom} onChange={e => setCompetencesList(competencesList.map(x => x.id === c.id ? { ...x, nom: e.target.value } : x))}
                      style={{ width: '100%', border: 'none', background: 'transparent', color: '#000', fontSize: '0.8rem' }} />
                  </td>
                  <td style={{ padding: '6px', textAlign: 'center' }}>
                    <button onClick={() => setCompetencesList(competencesList.filter(x => x.id !== c.id))} style={{ background: 'none', border: 'none', cursor: 'pointer', fontSize: '0.85rem' }}>❌</button>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>

        {/* Liste des ressources */}
        <div>
          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '0.8rem' }}>
            <h3 style={{ color: '#469b6f', margin: 0, fontSize: '0.95rem' }}>👥 Ressources configurées ({ressourcesList.length})</h3>
            <button onClick={() => setRessourcesList([...ressourcesList, { id: Date.now(), nom: 'Nouvelle ressource', fonction: '', niveau_cible: 3 }])} style={{ padding: '4px 10px', background: '#469b6f', color: '#fff', border: 'none', borderRadius: '4px', cursor: 'pointer', fontSize: '0.8rem' }}>+ Ajouter</button>
          </div>
          <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '0.8rem', border: '1px solid #d1fae5' }}>
            <thead style={{ background: '#cbe4d5', color: '#166534' }}>
              <tr>
                <th style={{ padding: '8px', textAlign: 'left', border: '1px solid #fff' }}>Nom</th>
                <th style={{ padding: '8px', textAlign: 'left', border: '1px solid #fff' }}>Fonction</th>
                <th style={{ padding: '8px', textAlign: 'center', border: '1px solid #fff' }}>Niv. cible</th>
                <th style={{ padding: '8px', textAlign: 'center', border: '1px solid #fff' }}>Action</th>
              </tr>
            </thead>
            <tbody>
              {ressourcesList.map((r) => (
                <tr key={r.id} style={{ borderBottom: '1px solid #f0fdf4' }}>
                  <td style={{ padding: '5px 8px' }}>
                    <input value={r.nom} onChange={e => setRessourcesList(ressourcesList.map(x => x.id === r.id ? { ...x, nom: e.target.value } : x))}
                      style={{ width: '100%', border: 'none', background: 'transparent', color: '#000', fontSize: '0.8rem' }} />
                  </td>
                  <td style={{ padding: '5px 8px' }}>
                    <input value={r.fonction} onChange={e => setRessourcesList(ressourcesList.map(x => x.id === r.id ? { ...x, fonction: e.target.value } : x))}
                      style={{ width: '100%', border: 'none', background: 'transparent', color: '#777', fontSize: '0.78rem' }} />
                  </td>
                  <td style={{ padding: '5px 8px', textAlign: 'center' }}>
                    <select value={r.niveau_cible} onChange={e => setRessourcesList(ressourcesList.map(x => x.id === r.id ? { ...x, niveau_cible: parseInt(e.target.value) } : x))}
                      style={{ border: '1px solid #ddd', borderRadius: '3px', padding: '2px', fontSize: '0.78rem' }}>
                      {[1, 2, 3, 4].map(n => <option key={n} value={n}>{n}</option>)}
                    </select>
                  </td>
                  <td style={{ padding: '5px', textAlign: 'center' }}>
                    <button onClick={() => setRessourcesList(ressourcesList.filter(x => x.id !== r.id))} style={{ background: 'none', border: 'none', cursor: 'pointer', fontSize: '0.85rem' }}>❌</button>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
};

// ─── RACIView ─────────────────────────────────────────────────────
// ─── RACI STAKEHOLDERS & VIEW ──────────────────────────────────────
const RACI_STAKEHOLDERS = [
  { group: 'Équipe Management', members: ['Sponsor', 'Changement', 'Informatique', 'Chef de projet', 'Intégration'] },
  { group: 'Équipe Projet', members: ['Coord. Finances', 'Coord. Ventes', 'Coord. Achats', 'Coord. Logistique', 'Coord. Production', 'Team Finances'] },
  { group: 'Équipe Support', members: ['Team Ventes', 'Team Achats', 'Team Logistique', 'Support Admin', 'Expert Master Data', 'Utilisateurs clés', 'Formateur'] }
];

const ALL_MEMBERS = RACI_STAKEHOLDERS.flatMap(g => g.members);

interface RACIRow {
  id: string;
  task: string;
  bold?: boolean;
  roles: Record<string, string>;
}

const RACI_ROWS: RACIRow[] = [
  { id: '1', task: 'INITIATION', bold: true, roles: {} },
  { id: '1.1', task: "Etude d'opportunité", roles: { 'Sponsor': 'R', 'Changement': 'R', 'Chef de projet': 'A', 'Coord. Finances': 'R', 'Coord. Achats': 'A', 'Coord. Logistique': 'C', 'Coord. Production': 'C', 'Team Finances': 'I', 'Team Ventes': 'I', 'Team Achats': 'R', 'Team Logistique': 'C', 'Support Admin': 'I', 'Expert Master Data': 'C' } },
  { id: '1.2', task: 'Charte Projet', roles: { 'Sponsor': 'A', 'Changement': 'R', 'Chef de projet': 'C', 'Intégration': 'IC', 'Coord. Ventes': 'CIR', 'Coord. Production': 'R', 'Team Finances': 'I', 'Team Ventes': 'R', 'Team Achats': 'I', 'Team Logistique': 'I', 'Utilisateurs clés': 'R' } },
  { id: '2', task: 'PLANIFICATION', bold: true, roles: {} },
  { id: '2.1', task: 'Cahier des Charges', roles: { 'Sponsor': 'R', 'Chef de projet': 'R', 'Informatique': 'C', 'Intégration': 'C', 'Coord. Ventes': 'ICR', 'Coord. Logistique': 'A', 'Team Finances': 'C', 'Team Ventes': 'C', 'Team Achats': 'C', 'Team Logistique': 'C', 'Support Admin': 'I', 'Expert Master Data': 'C', 'Utilisateurs clés': 'C', 'Formateur': 'R' } },
  { id: '2.2', task: 'Plan détaillé', roles: { 'Sponsor': 'A', 'Changement': 'C', 'Informatique': 'C', 'Chef de projet': 'I', 'Intégration': 'C', 'Coord. Achats': 'C', 'Coord. Logistique': 'C', 'Coord. Production': 'C', 'Team Finances': 'C', 'Team Ventes': 'I', 'Team Achats': 'I', 'Team Logistique': 'I', 'Support Admin': 'I', 'Expert Master Data': 'I', 'Utilisateurs clés': 'I', 'Formateur': 'I' } },
  { id: '3', task: 'RÉALISATION', bold: true, roles: {} },
  { id: '3.1', task: 'Registre des risques et problèmes', roles: { 'Chef de projet': 'R', 'Intégration': 'A', 'Changement': 'C', 'Informatique': 'C', 'Coord. Achats': 'C', 'Coord. Logistique': 'R', 'Coord. Production': 'I', 'Team Logistique': 'I', 'Expert Master Data': 'C' } },
  { id: '3.2', task: 'Plan de Communication', roles: { 'Changement': 'R', 'Informatique': 'A', 'Chef de projet': 'C', 'Coord. Ventes': 'C', 'Coord. Achats': 'C', 'Coord. Logistique': 'C', 'Coord. Production': 'C', 'Team Finances': 'I', 'Team Ventes': 'I', 'Team Achats': 'I', 'Team Logistique': 'I', 'Support Admin': 'I', 'Utilisateurs clés': 'I', 'Formateur': 'C' } },
  { id: '3.3', task: 'Mise en place Finances', roles: { 'Sponsor': 'I', 'Informatique': 'I', 'Chef de projet': 'C', 'Intégration': 'C', 'Coord. Finances': 'R', 'Coord. Ventes': 'I', 'Coord. Achats': 'I', 'Coord. Logistique': 'I', 'Coord. Production': 'I', 'Team Finances': 'C', 'Team Ventes': 'I', 'Team Achats': 'I', 'Team Logistique': 'I', 'Expert Master Data': 'I', 'Utilisateurs clés': 'I' } },
  { id: '3.4', task: 'Mise en place Ventes', roles: { 'Sponsor': 'I', 'Informatique': 'I', 'Chef de projet': 'A', 'Intégration': 'C', 'Coord. Ventes': 'R', 'Coord. Logistique': 'I', 'Coord. Production': 'C', 'Team Ventes': 'C', 'Team Achats': 'I', 'Team Logistique': 'I', 'Utilisateurs clés': 'C' } },
  { id: '3.5', task: 'Mise en place Achats', roles: { 'Sponsor': 'I', 'Changement': 'R', 'Informatique': 'I', 'Intégration': 'C', 'Coord. Finances': 'A', 'Coord. Achats': 'R', 'Coord. Logistique': 'I', 'Coord. Production': 'I', 'Team Achats': 'C', 'Team Logistique': 'I', 'Expert Master Data': 'I', 'Utilisateurs clés': 'I' } },
  { id: '3.6', task: 'Mise en place Production', roles: { 'Sponsor': 'I', 'Chef de projet': 'I', 'Coord. Logistique': 'I', 'Coord. Production': 'R', 'Team Finances': 'I', 'Team Achats': 'C', 'Team Logistique': 'I', 'Support Admin': 'A', 'Expert Master Data': 'C' } },
  { id: '3.7', task: 'Mise en place Logistique', roles: { 'Informatique': 'A', 'Intégration': 'C', 'Coord. Logistique': 'R', 'Team Logistique': 'R' } },
  { id: '3.8', task: 'Tests et Validation', roles: { 'Sponsor': 'I', 'Informatique': 'I', 'Chef de projet': 'C', 'Intégration': 'A', 'Coord. Finances': 'C', 'Coord. Ventes': 'C', 'Coord. Achats': 'C', 'Coord. Logistique': 'C', 'Coord. Production': 'C', 'Team Finances': 'I', 'Team Ventes': 'I', 'Team Achats': 'I', 'Team Logistique': 'I', 'Expert Master Data': 'I', 'Utilisateurs clés': 'R' } },
  { id: '4', task: 'CONTRÔLE ET SUIVI', bold: true, roles: {} },
  { id: '4.1', task: "Tests d'intégration", roles: { 'Sponsor': 'I', 'Informatique': 'I', 'Chef de projet': 'C', 'Coord. Finances': 'C', 'Coord. Ventes': 'C', 'Coord. Achats': 'A', 'Coord. Logistique': 'C', 'Coord. Production': 'C', 'Team Finances': 'I', 'Team Ventes': 'I', 'Team Achats': 'I', 'Team Logistique': 'I', 'Utilisateurs clés': 'R' } },
  { id: '4.2', task: 'Tests de Consolidation', roles: { 'Sponsor': 'I', 'Informatique': 'I', 'Chef de projet': 'C', 'Coord. Ventes': 'A', 'Coord. Logistique': 'C', 'Coord. Production': 'C', 'Team Finances': 'I', 'Team Ventes': 'I', 'Team Achats': 'I', 'Team Logistique': 'I', 'Utilisateurs clés': 'R' } },
  { id: '4.3', task: 'Formation utilisateurs CMRL', roles: { 'Sponsor': 'I', 'Chef de projet': 'C', 'Intégration': 'C', 'Coord. Finances': 'I', 'Coord. Ventes': 'A', 'Coord. Achats': 'I', 'Coord. Logistique': 'I', 'Coord. Production': 'I', 'Team Finances': 'I', 'Team Ventes': 'I', 'Team Achats': 'I', 'Team Logistique': 'I', 'Utilisateurs clés': 'R', 'Formateur': 'R' } },
  { id: '4.4', task: 'Déploiement et Basculement', roles: { 'Intégration': 'R', 'Coord. Finances': 'C', 'Coord. Ventes': 'C', 'Coord. Achats': 'C', 'Coord. Logistique': 'C', 'Coord. Production': 'C', 'Team Ventes': 'A', 'Team Logistique': 'I', 'Utilisateurs clés': 'I' } },
  { id: '4.5', task: 'Validation et confirmation', roles: { 'Chef de projet': 'R', 'Coord. Logistique': 'A', 'Team Logistique': 'I' } },
  { id: '5', task: 'CLÔTURE', bold: true, roles: {} },
  { id: '5.1', task: 'Liste des corrections', roles: { 'Sponsor': 'C', 'Changement': 'C', 'Informatique': 'C', 'Chef de projet': 'A', 'Coord. Finances': 'R', 'Coord. Ventes': 'R', 'Coord. Achats': 'R', 'Coord. Logistique': 'R', 'Coord. Production': 'R', 'Team Finances': 'C', 'Team Ventes': 'C', 'Team Achats': 'C', 'Team Logistique': 'C', 'Support Admin': 'C', 'Expert Master Data': 'C', 'Utilisateurs clés': 'C', 'Formateur': 'C' } },
  { id: '5.2', task: 'Leçons apprises', roles: { 'Sponsor': 'C', 'Changement': 'C', 'Chef de projet': 'R', 'Coord. Achats': 'C', 'Team Ventes': 'A', 'Utilisateurs clés': 'C' } },
  { id: '5.3', task: 'Décommissionnement', roles: { 'Informatique': 'R', 'Intégration': 'C', 'Coord. Logistique': 'IR', 'Team Achats': 'A' } },
  { id: '5.4', task: 'Réunion de clôture', roles: { 'Sponsor': 'A', 'Chef de projet': 'R', 'Intégration': 'R', 'Coord. Production': 'R', 'Team Achats': 'C', 'Support Admin': 'I' } },
];

const RACIView = ({ projectId, fileId, currentDossierId, page = 'option1' }: { projectId: string, fileId: string, currentDossierId: string | null, page?: string }) => {
  const dossierSubKey = currentDossierId ? `-${currentDossierId}` : "";
  const storageKey = `faab-raci-data-${projectId}-${fileId}${dossierSubKey}`;
  const stakeholdersKey = `faab-raci-stakeholders-${projectId}-${fileId}${dossierSubKey}`;

  const defaultStakeholders = [
    { group: 'Équipe Management', members: ['Sponsor', 'Changement', 'Informatique', 'Chef de projet', 'Intégration'] },
    { group: 'Équipe Projet', members: ['Coord. Finances', 'Coord. Ventes', 'Coord. Achats', 'Coord. Logistique'] }
  ];

  const defaultRows: RACIRow[] = [
    { id: '1', task: 'INITIATION', bold: true, roles: {} },
    { id: '1.1', task: "Etude d'opportunité", roles: { 'Sponsor': 'R', 'Changement': 'R', 'Chef de projet': 'A' } },
    { id: '2', task: 'PLANIFICATION', bold: true, roles: {} },
    { id: '2.1', task: 'Cahier des Charges', roles: { 'Sponsor': 'R', 'Chef de projet': 'R' } },
  ];

  const [stakeholders, setStakeholders] = useState(() => {
    const saved = localStorage.getItem(stakeholdersKey);
    return saved ? JSON.parse(saved) : defaultStakeholders;
  });

  const [rows, setRows] = useState(() => {
    const saved = localStorage.getItem(storageKey);
    return saved ? JSON.parse(saved) : defaultRows;
  });

  useEffect(() => {
    localStorage.setItem(stakeholdersKey, JSON.stringify(stakeholders));
  }, [stakeholders, stakeholdersKey]);

  useEffect(() => {
    localStorage.setItem(storageKey, JSON.stringify(rows));
  }, [rows, storageKey]);

  const allMembers = stakeholders.flatMap((g: any) => g.members);

  const updateRole = (rowId: string, member: string, val: string) => {
    setRows(rows.map((r: RACIRow) => r.id === rowId ? { ...r, roles: { ...r.roles, [member]: val } } : r));
  };

  const getCellStyles = (val: string) => {
    if (!val) return {};
    const base = { fontWeight: 900 as const, textAlign: 'center' as const, fontSize: '0.8rem', borderRadius: '4px' };
    if (val.includes('R')) return { ...base, color: '#16a34a', background: 'rgba(22, 163, 74, 0.1)' };
    if (val.includes('A')) return { ...base, color: '#2563eb', background: 'rgba(37, 99, 235, 0.1)' };
    if (val.includes('C')) return { ...base, color: '#ca8a04', background: 'rgba(202, 138, 4, 0.1)' };
    if (val.includes('I')) return { ...base, color: '#7c3aed', background: 'rgba(124, 58, 237, 0.1)' };
    return base;
  };

  return (
    <div style={{ flex: 1, overflow: 'auto', background: 'var(--main-bg)', color: 'var(--main-text)', padding: '2rem' }}>
      <div style={{ maxWidth: '1600px', margin: '0 auto' }}>
        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '2rem', background: 'var(--main-surface)', padding: '1.5rem', borderRadius: '16px', border: '1px solid var(--main-border)', boxShadow: '0 4px 12px rgba(0,0,0,0.05)' }}>
          <div style={{ display: 'flex', alignItems: 'center', gap: '1rem' }}>
            <div style={{ width: '40px', height: '40px', borderRadius: '12px', background: 'var(--table-primary)', display: 'flex', alignItems: 'center', justifyContent: 'center', color: '#fff', fontSize: '1.2rem' }}>📊</div>
            <h2 style={{ margin: 0, fontWeight: 900, color: 'var(--table-primary)' }}>Matrice RACI du Projet</h2>
          </div>
          <div style={{ display: 'flex', gap: '1rem' }}>
            <div style={{ padding: '6px 16px', background: 'rgba(22, 163, 74, 0.1)', color: '#16a34a', borderRadius: '8px', fontWeight: 800, fontSize: '0.75rem' }}>R: Réalisateur</div>
            <div style={{ padding: '6px 16px', background: 'rgba(37, 99, 235, 0.1)', color: '#2563eb', borderRadius: '8px', fontWeight: 800, fontSize: '0.75rem' }}>A: Approbateur</div>
            <div style={{ padding: '6px 16px', background: 'rgba(202, 138, 4, 0.1)', color: '#ca8a04', borderRadius: '8px', fontWeight: 800, fontSize: '0.75rem' }}>C: Consulté</div>
            <div style={{ padding: '6px 16px', background: 'rgba(124, 58, 237, 0.1)', color: '#7c3aed', borderRadius: '8px', fontWeight: 800, fontSize: '0.75rem' }}>I: Informé</div>
          </div>
        </div>

        <div style={{ background: 'var(--main-surface)', borderRadius: '24px', border: '1px solid var(--main-border)', overflow: 'hidden', boxShadow: '0 8px 32px rgba(0,0,0,0.05)' }}>
          <div style={{ overflowX: 'auto' }}>
            <table style={{ borderCollapse: 'collapse', width: 'max-content', minWidth: '100%' }}>
              <thead>
                <tr style={{ background: 'rgba(var(--table-primary-rgb), 0.05)' }}>
                  <th style={{ padding: '12px', borderRight: '1px solid var(--main-border)' }} colSpan={2}></th>
                  {stakeholders.map((g: any, i: number) => (
                    <th key={i} colSpan={g.members.length} style={{ padding: '10px', borderRight: '1px solid var(--main-border)', borderBottom: '1px solid var(--main-border)', fontSize: '0.7rem', fontWeight: 900, color: 'var(--table-primary)', textTransform: 'uppercase' }}>{g.group}</th>
                  ))}
                </tr>
                <tr style={{ background: 'var(--table-primary)', color: '#fff' }}>
                  <th style={{ padding: '12px', textAlign: 'left', width: '80px', borderRight: '1px solid rgba(255,255,255,0.2)' }}>ID</th>
                  <th style={{ padding: '12px', textAlign: 'left', width: '300px', borderRight: '1px solid rgba(255,255,255,0.2)' }}>Processus / Activités</th>
                  {allMembers.map((m: string, i: number) => (
                    <th key={i} style={{ padding: '10px', fontSize: '0.65rem', width: '80px', borderRight: i === allMembers.length - 1 ? 'none' : '1px solid rgba(255,255,255,0.2)' }}>{m}</th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {rows
                  .filter((r: RACIRow) => {
                    if (page === 'option1') return true;
                    if (r.bold) return true;
                    const values = Object.values(r.roles || {});
                    if (page === 'option2') return values.some(v => v.includes('R'));
                    if (page === 'option3') return values.some(v => v.includes('A'));
                    if (page === 'option4') return values.some(v => v.includes('C') || v.includes('I'));
                    return true;
                  })
                  .map((r: RACIRow) => (
                    <tr key={r.id} style={{ borderBottom: '1px solid var(--main-border)', background: r.bold ? 'rgba(var(--table-primary-rgb), 0.02)' : 'transparent' }}>
                      <td style={{ padding: '10px', fontSize: '0.75rem', fontWeight: 900, color: 'var(--text-muted)', borderRight: '1px solid var(--main-border)' }}>{r.id}</td>
                      <td style={{ padding: '10px', fontWeight: r.bold ? 900 : 700, fontSize: r.bold ? '0.85rem' : '0.8rem', borderRight: '1px solid var(--main-border)' }}>{r.task}</td>
                      {allMembers.map((m: string, j: number) => (
                        <td key={j} style={{ padding: '4px', borderRight: j === allMembers.length - 1 ? 'none' : '1px solid var(--main-border)', ...getCellStyles(r.roles?.[m] || '') }}>
                          {!r.bold && (
                            <input
                              value={r.roles?.[m] || ''}
                              onChange={(e) => updateRole(r.id, m, e.target.value.toUpperCase())}
                              style={{ width: '100%', border: 'none', background: 'transparent', textAlign: 'center', fontWeight: 900, color: 'inherit', height: '25px' }}
                              placeholder="-"
                            />
                          )}
                        </td>
                      ))}
                    </tr>
                  ))}
              </tbody>
            </table>
          </div>
          <div style={{ padding: '1.5rem', display: 'flex', gap: '1rem' }}>
            <button onClick={() => setRows([...rows, { id: `${rows.length + 1}`, task: 'Nouvelle tâche', roles: {} }])} style={{ padding: '10px 24px', background: 'var(--table-primary)', color: '#fff', border: 'none', borderRadius: '12px', fontWeight: 800, cursor: 'pointer' }}>+ Ajouter une ligne</button>
          </div>
        </div>
      </div>
    </div>
  );
};


// ─── PlanChargeView ───────────────────────────────────────────────
// ─── PlanChargeView ───────────────────────────────────────────────
const PlanChargeView = ({ projectId, fileId, currentDossierId, page = 'option2', onPageChange }: {
  projectId: string;
  fileId: string;
  currentDossierId: string | null;
  page?: string;
  onPageChange?: (p: string) => void;
}) => {
  const dossierSubKey = currentDossierId ? `-${currentDossierId}` : "";
  const storageKeyTasks = `faab-plan-charge-tasks-${projectId}-${fileId}${dossierSubKey}`;
  const storageKeyResources = `faab-plan-charge-resources-${projectId}-${fileId}${dossierSubKey}`;

  const defaultTaskData = [
    { id: 'de toutes les activités', task: 'MODALITES DE FONCTIONNEMENT', res: '', charge: 2596, prog: 15.8, isHeader: true },
    { id: '1.0', task: 'Planification Macro', res: 'Chantal Dupont', charge: 12, prog: 100 },
    { id: '1.1', task: 'Conception SAP', res: 'Jacques Meunier', charge: 45, prog: 20 },
    { id: '1.2', task: 'Paramétrage FI/CO', res: 'Monique Davant', charge: 60, prog: 10 },
  ];

  const defaultResourcesData = [
    { id: 1, name: 'Chantal Dupont', role: 'Chef de Projet', email: 'DupontC@confassis.com', dispo: '100%', de: '29/03/2021', a: '31/01/2022', dispoJ: '201,00', planJ: '109', charge: '54%', conges: 20 },
    { id: 2, name: 'Olivier Robespierre', role: 'Sponsor', email: 'RobespierreO@confassis.com', dispo: '20%', de: '29/03/2021', a: '31/01/2022', dispoJ: '43,20', planJ: '0', charge: '0%', conges: 5 },
    { id: 3, name: 'Jacques Meunier', role: 'Responsable Intégration', email: 'MeunierJ@confassis.com', dispo: '100%', de: '29/03/2021', a: '31/01/2022', dispoJ: '211,00', planJ: '189', charge: '90%', conges: 10 },
    { id: 4, name: 'Monique Davant', role: 'Coordinatrice Finances', email: 'DavantM@confassis.com', dispo: '100%', de: '12/04/2021', a: '15/01/2022', dispoJ: '190,00', planJ: '199', charge: '105%', conges: 10 },
  ];

  const [taskData, setTaskData] = useState(() => {
    const saved = localStorage.getItem(storageKeyTasks);
    return saved ? JSON.parse(saved) : defaultTaskData;
  });

  const [resourcesData, setResourcesData] = useState(() => {
    const saved = localStorage.getItem(storageKeyResources);
    return saved ? JSON.parse(saved) : defaultResourcesData;
  });

  useEffect(() => {
    localStorage.setItem(storageKeyTasks, JSON.stringify(taskData));
  }, [taskData, storageKeyTasks]);

  useEffect(() => {
    localStorage.setItem(storageKeyResources, JSON.stringify(resourcesData));
  }, [resourcesData, storageKeyResources]);

  const navTabStyle = (active: boolean): React.CSSProperties => ({
    background: active ? 'var(--table-primary)' : 'var(--main-surface)',
    color: active ? '#fff' : 'var(--text-muted)',
    padding: '12px 24px',
    borderRadius: '12px',
    fontWeight: 800,
    fontSize: '0.85rem',
    cursor: 'pointer',
    border: active ? 'none' : '1px solid var(--main-border)',
    boxShadow: active ? '0 4px 12px rgba(var(--table-primary-rgb), 0.3)' : 'none',
    display: 'flex',
    alignItems: 'center',
    gap: '8px'
  });

  return (
    <div style={{ flex: 1, display: 'flex', flexDirection: 'column', background: 'var(--main-bg)', color: 'var(--main-text)', padding: '2rem', overflow: 'auto' }}>
      <div style={{ maxWidth: '1400px', margin: '0 auto', width: '100%' }}>
        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '2.5rem' }}>
          <div style={{ display: 'flex', alignItems: 'center', gap: '1rem' }}>
            <div style={{ width: '45px', height: '45px', borderRadius: '14px', background: 'var(--table-primary)', display: 'flex', alignItems: 'center', justifyContent: 'center', color: '#fff', fontSize: '1.5rem' }}>⌛</div>
            <h2 style={{ margin: 0, fontWeight: 900, color: 'var(--table-primary)', fontSize: '1.8rem' }}>Plan de Charge des Ressources</h2>
          </div>
          <div style={{ display: 'flex', gap: '10px' }}>
            <div onClick={() => onPageChange?.('option1')} style={navTabStyle(page === 'option1')}>📖 Guide</div>
            <div onClick={() => onPageChange?.('option2')} style={navTabStyle(page === 'option2')}>📊 Plan de charge</div>
            <div onClick={() => onPageChange?.('option3')} style={navTabStyle(page === 'option3')}>👥 Ressources</div>
          </div>
        </div>

        {page === 'option1' && (
          <div style={{ background: 'var(--main-surface)', padding: '3rem', borderRadius: '32px', border: '1px solid var(--main-border)', boxShadow: '0 12px 48px rgba(0,0,0,0.05)' }}>
            <p style={{ fontSize: '1.1rem', lineHeight: 1.8, color: 'var(--text-muted)' }}>
              Ce module vous permet de surveiller la charge de travail de chaque membre de l'équipe.
            </p>
          </div>
        )}

        {page === 'option2' && (
          <div style={{ background: 'var(--main-surface)', borderRadius: '24px', border: '1px solid var(--main-border)', overflow: 'hidden' }}>
            <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '0.85rem' }}>
              <thead style={{ background: 'var(--table-primary)', color: '#fff' }}>
                <tr>
                  <th style={{ padding: '15px', textAlign: 'left' }}>Tâche / Activité</th>
                  <th style={{ padding: '15px', textAlign: 'left' }}>Ressource</th>
                  <th style={{ padding: '15px', textAlign: 'center' }}>Charge (j)</th>
                  <th style={{ padding: '15px', textAlign: 'center' }}>Progression</th>
                </tr>
              </thead>
              <tbody>
                {taskData.map((t: any, idx: number) => (
                  <tr key={idx} style={{ borderBottom: '1px solid var(--main-border)', background: t.isHeader ? 'rgba(var(--table-primary-rgb), 0.05)' : 'transparent' }}>
                    <td style={{ padding: '12px 15px', fontWeight: t.isHeader ? 900 : 500 }}>{t.task}</td>
                    <td style={{ padding: '12px 15px' }}>{t.res}</td>
                    <td style={{ padding: '12px 15px', textAlign: 'center' }}>{t.charge}</td>
                    <td style={{ padding: '12px 15px', textAlign: 'center' }}>{t.prog}%</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}

        {page === 'option3' && (
          <div style={{ background: 'var(--main-surface)', borderRadius: '24px', border: '1px solid var(--main-border)', overflow: 'hidden' }}>
            <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '0.85rem' }}>
              <thead style={{ background: 'var(--table-primary)', color: '#fff' }}>
                <tr>
                  <th style={{ padding: '15px', textAlign: 'left' }}>Nom</th>
                  <th style={{ padding: '15px', textAlign: 'left' }}>Rôle</th>
                  <th style={{ padding: '15px', textAlign: 'center' }}>Dispo J</th>
                  <th style={{ padding: '15px', textAlign: 'center' }}>Planifié J</th>
                  <th style={{ padding: '15px', textAlign: 'center' }}>Charge %</th>
                </tr>
              </thead>
              <tbody>
                {resourcesData.map((r: any, idx: number) => (
                  <tr key={idx} style={{ borderBottom: '1px solid var(--main-border)' }}>
                    <td style={{ padding: '12px 15px', fontWeight: 800 }}>{r.name}</td>
                    <td style={{ padding: '12px 15px' }}>{r.role}</td>
                    <td style={{ padding: '12px 15px', textAlign: 'center' }}>{r.dispoJ}</td>
                    <td style={{ padding: '12px 15px', textAlign: 'center' }}>{r.planJ}</td>
                    <td style={{ padding: '12px 15px', textAlign: 'center' }}>{r.charge}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}
      </div>
    </div>
  );
};


const ExecutionHeader = ({ title, metadata, light = false }: { title: string; metadata: any; light?: boolean }) => (
  <div style={{
    background: 'var(--main-surface)',
    padding: '1.5rem 2.5rem',
    borderBottom: '1px solid var(--main-border)',
    color: 'var(--main-text)',
    backdropFilter: 'blur(10px)',
    position: 'sticky',
    top: 0,
    zIndex: 10
  }}>
    <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
      <div style={{ display: 'flex', gap: '1.25rem', alignItems: 'center' }}>
        <div style={{
          background: 'linear-gradient(135deg, #10b981, #059669)',
          padding: '12px',
          borderRadius: '12px',
          boxShadow: '0 4px 12px rgba(16, 185, 129, 0.2)'
        }}>
          <svg width="24" height="24" viewBox="0 0 24 24" fill="white"><path d="M12 2L2 7l10 5 10-5-10-5zM2 17l10 5 10-5M2 12l10 5 10-5" /></svg>
        </div>
        <div>
          <h1 style={{ margin: 0, fontSize: '1.5rem', fontWeight: 800, letterSpacing: '-0.02em', color: 'var(--main-text)' }}>{title}</h1>
        </div>
      </div>
      <div style={{ textAlign: 'right' }}>
      </div>
    </div>
  </div>
);


const SuiviLivrablesView = ({ page, projectId, fileId, currentDossierId }: { page: 'livrables' | 'parametrage', projectId: string, fileId: string, currentDossierId: string | null }) => {
  const metadata = { projectName: 'SAP30 / CMRL', manager: 'Mme Dupont', version: 'V2.3', lastUpdate: '28/11/2022' };
  const dossierSubKey = currentDossierId ? `-${currentDossierId}` : "";
  const storageKey = `faab-livrables-data-${projectId}-${fileId}${dossierSubKey}`;

  const [livrables, setLivrables] = useState(() => {
    const saved = localStorage.getItem(storageKey);
    if (saved) return JSON.parse(saved);
    return [
      { id: 1, name: "Étude d'opportunité", type: 'Support décisionnel', resp: 'Chantal / Oliver', desc: "Justification du projet et sa place dans la stratégie d'entreprise", etat: 'Livré', prev: '9/4/2022', real: '10/4/2022', val: 'Comité de direction' },
      { id: 2, name: 'Charte Projet', type: 'Initiation', resp: 'Chantal / Oliver', desc: 'Charte de projet détaillée y inclus délais et coûts estimés', etat: 'Livré', prev: '29/4/2022', real: '28/4/2022', val: 'Comité de direction' },
      { id: 3, name: 'Macro planning', type: 'Initiation', resp: 'Chantal / Oliver', desc: 'Estimation des phases principales', etat: 'Livré', prev: '29/4/2022', real: '28/4/2022', val: 'Comité de direction' },
      { id: 4, name: 'Support réunion de lancement', type: 'Planification', resp: 'Chantal / Oliver', desc: "Présentation du projet destinée à l'équipe dès que la charte sera approuvée", etat: 'Livré', prev: '29/4/2022', real: '29/4/2022', val: 'Sponsor' },
      { id: 5, name: 'Expression des besoins', type: 'Planification', resp: 'Chantal / Oliver', desc: 'Résumé des besoins exprimés lors des réunions de découverte', etat: 'En cours', prev: '29/6/2022', real: '-', val: 'Sponsor' },
      { id: 6, name: 'Adéquation exigences', type: 'Planification', resp: 'Chantal / Oliver', desc: 'Matrice d’adéquation exigences / fonctionnalités', etat: 'En attente', prev: '29/6/2022', real: '-', val: 'Sponsor' },
      { id: 7, name: 'Cahier des charges', type: 'Planification', resp: 'Chantal / Oliver', desc: 'Spécifications détaillées', etat: 'En cours', prev: '30/6/2022', real: '-', val: 'Comité de direction' },
      { id: 8, name: 'Plan de projet', type: 'Planification', resp: 'Chantal / Oliver', desc: 'Plan de management consolidé', etat: 'Livré', prev: '30/6/2022', real: '30/6/2022', val: 'Comité de direction' },
      { id: 9, name: 'Plan de charge', type: 'Planification', resp: 'Chantal / Oliver', desc: 'Allocation des ressources j-h', etat: 'Livré', prev: '29/6/2022', real: '28/6/2022', val: 'Comité de direction' },
      { id: 10, name: 'Registre des problèmes et risques', type: 'Planification', resp: 'Chantal / Oliver', desc: 'Identification et cotation des risques', etat: 'En cours', prev: '29/6/2022', real: '-', val: 'Sponsor' },
      { id: 11, name: 'Registre des parties prenantes', type: 'Planification', resp: 'Chantal / Oliver', desc: 'Cartographie acteurs clés', etat: 'Livré', prev: '29/6/2022', real: '25/6/2022', val: 'Sponsor' },
      { id: 12, name: 'Matrice des compétences', type: 'Planification', resp: 'Chantal / Oliver', desc: 'Grille des savoir-faire requis', etat: 'Livré', prev: '29/6/2022', real: '27/6/2022', val: 'Sponsor' },
      { id: 13, name: 'Plan de communication', type: 'Planification', resp: 'Chantal / Oliver', desc: 'Planning des annonces et livrables de com', etat: 'En cours', prev: '29/6/2022', real: '-', val: 'Comité de direction' },
      { id: 14, name: 'Suivi des actions', type: 'Planification', resp: 'Chantal / Oliver', desc: 'Journal opérationnel', etat: 'En cours', prev: '29/6/2022', real: '-', val: 'Sponsor' },
      { id: 15, name: 'Suivi des livrables', type: 'Planification', resp: 'Chantal / Oliver', desc: 'Registre de validation', etat: 'En cours', prev: '29/6/2022', real: '-', val: 'Sponsor' },
      { id: 16, name: 'Registre des demandes de changement', type: 'Réalisation', resp: 'Chantal / Coordinateurs', desc: 'Gestion des évolutions du périmètre', etat: 'En attente', prev: '15/7/2022', real: '-', val: 'Sponsor / Communication' },
      { id: 17, name: 'Tableau de bord', type: 'Rapports', resp: 'Chantal', desc: 'Tableau de bord de suivi consolidé', etat: 'Livré', prev: '15/7/2022', real: '15/7/2022', val: 'Sponsor / Communication' },
      { id: 18, name: 'Rapport de suivi de projet', type: 'Rapports', resp: 'Chantal', desc: 'Synthèse mensuelle pour le comité', etat: 'En cours', prev: '15/7/2022', real: '-', val: 'Sponsor / Communication' },
      { id: 19, name: 'Rapport Flash', type: 'Rapports', resp: 'Chantal', desc: 'Fiche météo hebdomadaire', etat: 'Livré', prev: '15/7/2022', real: '14/7/2022', val: 'Sponsor / Communication' }
    ];
  });

  useEffect(() => {
    localStorage.setItem(storageKey, JSON.stringify(livrables));
  }, [livrables]);

  const addLivrable = () => {
    const newId = livrables.length > 0 ? Math.max(...livrables.map((l: any) => l.id)) + 1 : 1;
    setLivrables([...livrables, { id: newId, name: 'Nouveau Livrable', type: 'Doc', resp: 'Équipe', desc: 'Description', etat: 'En attente', prev: '-', real: '-', val: '-' }]);
  };

  const deleteLivrable = (id: number) => {
    setLivrables(livrables.filter((l: any) => l.id !== id));
  };

  const handleChange = (id: number, field: string, val: any) => {
    setLivrables(livrables.map((l: any) => l.id === id ? { ...l, [field]: val } : l));
  };

  const totalCount = livrables.length;
  const livresCount = livrables.filter((l: any) => l.etat === 'Livré').length;
  const enCoursCount = livrables.filter((l: any) => l.etat === 'En cours').length;
  const enAttenteCount = livrables.filter((l: any) => l.etat === 'En attente' || l.etat === 'Non commencé').length;
  const pctAvancement = totalCount > 0 ? Math.round((livresCount / totalCount) * 100) : 0;

  if (page === 'parametrage') {
    return (
      <div style={{ flex: 1, display: 'flex', flexDirection: 'column', background: 'var(--main-bg)' }}>
        <ExecutionHeader title="23 — Paramétrage des Livrables" metadata={metadata} />
        <div style={{ padding: '2.5rem' }}>
          <div style={{ background: 'var(--main-surface)', borderRadius: '16px', border: '1px solid var(--main-border)', padding: '2.5rem', boxShadow: 'var(--card-shadow)' }}>
            <h3 style={{ color: 'var(--excel-color)', marginBottom: '2rem', display: 'flex', alignItems: 'center', gap: '12px', fontSize: '1.25rem' }}>
              <span style={{ fontSize: '1.5rem' }}>⚙️</span> Configuration des Workflows
            </h3>
            <div style={{ overflow: 'hidden', borderRadius: '12px', border: '1px solid var(--main-border)' }}>
              <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '0.9rem' }}>
                <thead style={{ background: 'var(--main-bg)' }}>
                  <tr style={{ color: 'var(--text-muted)' }}>
                    <th style={{ padding: '18px 24px', textAlign: 'left', fontWeight: 700 }}>Type de document</th>
                    <th style={{ padding: '18px 24px', fontWeight: 700 }}>Délai standard</th>
                    <th style={{ padding: '18px 24px', textAlign: 'left', fontWeight: 700 }}>Circuit de validation</th>
                  </tr>
                </thead>
                <tbody>
                  <tr style={{ borderBottom: '1px solid var(--main-border)' }}>
                    <td style={{ padding: '18px 24px', fontWeight: 600, color: 'var(--main-text)' }}>Support décisionnel</td>
                    <td style={{ padding: '18px 24px', textAlign: 'center', color: 'var(--main-text)' }}>15 jours ouvrés</td>
                    <td style={{ padding: '18px 24px', color: 'var(--main-text)' }}>Direction Générale / COMEX</td>
                  </tr>
                  <tr>
                    <td style={{ padding: '18px 24px', fontWeight: 600, color: 'var(--main-text)' }}>Cahier des charges</td>
                    <td style={{ padding: '18px 24px', textAlign: 'center', color: 'var(--main-text)' }}>30 jours calendaires</td>
                    <td style={{ padding: '18px 24px', color: 'var(--main-text)' }}>Sponsor / Équipe Métier</td>
                  </tr>
                </tbody>
              </table>
            </div>
          </div>
        </div>
      </div>
    );
  }

  return (
    <div style={{ flex: 1, display: 'flex', flexDirection: 'column', background: 'var(--main-bg)', color: 'var(--main-text)' }}>
      <ExecutionHeader title="23 — Suivi des livrables" metadata={metadata} />
      
      {/* Dynamic KPI Row */}
      <div style={{ padding: '1.5rem 2.5rem 0', display: 'grid', gridTemplateColumns: 'repeat(4, 1fr)', gap: '1.25rem' }}>
        <div style={{ background: 'var(--main-surface)', padding: '1.25rem', borderRadius: '14px', border: '1px solid var(--main-border)', borderLeft: '4px solid #10b981', boxShadow: 'var(--card-shadow)' }}>
          <div style={{ fontSize: '0.75rem', fontWeight: 800, color: 'var(--text-muted)', textTransform: 'uppercase' }}>Total Livrables</div>
          <div style={{ fontSize: '1.8rem', fontWeight: 950, color: '#10b981', marginTop: '4px' }}>{totalCount}</div>
          <div style={{ fontSize: '0.75rem', color: 'var(--text-muted)' }}>Avancement : <b>{pctAvancement}%</b></div>
        </div>
        <div style={{ background: 'var(--main-surface)', padding: '1.25rem', borderRadius: '14px', border: '1px solid var(--main-border)', borderLeft: '4px solid #10b981', boxShadow: 'var(--card-shadow)' }}>
          <div style={{ fontSize: '0.75rem', fontWeight: 800, color: 'var(--text-muted)', textTransform: 'uppercase' }}>Livrés ✅</div>
          <div style={{ fontSize: '1.8rem', fontWeight: 950, color: '#10b981', marginTop: '4px' }}>{livresCount}</div>
          <div style={{ fontSize: '0.75rem', color: '#10b981', fontWeight: 700 }}>Conformes aux attentes</div>
        </div>
        <div style={{ background: 'var(--main-surface)', padding: '1.25rem', borderRadius: '14px', border: '1px solid var(--main-border)', borderLeft: '4px solid #3b82f6', boxShadow: 'var(--card-shadow)' }}>
          <div style={{ fontSize: '0.75rem', fontWeight: 800, color: 'var(--text-muted)', textTransform: 'uppercase' }}>En Cours 🔵</div>
          <div style={{ fontSize: '1.8rem', fontWeight: 950, color: '#3b82f6', marginTop: '4px' }}>{enCoursCount}</div>
          <div style={{ fontSize: '0.75rem', color: '#3b82f6', fontWeight: 700 }}>Production active</div>
        </div>
        <div style={{ background: 'var(--main-surface)', padding: '1.25rem', borderRadius: '14px', border: '1px solid var(--main-border)', borderLeft: '4px solid #f59e0b', boxShadow: 'var(--card-shadow)' }}>
          <div style={{ fontSize: '0.75rem', fontWeight: 800, color: 'var(--text-muted)', textTransform: 'uppercase' }}>En Attente ⏳</div>
          <div style={{ fontSize: '1.8rem', fontWeight: 950, color: '#f59e0b', marginTop: '4px' }}>{enAttenteCount}</div>
          <div style={{ fontSize: '0.75rem', color: '#f59e0b', fontWeight: 700 }}>Prochaines étapes</div>
        </div>
      </div>

      <div style={{ padding: '1.25rem 2.5rem', background: 'transparent', display: 'flex', gap: '15px', alignItems: 'center' }}>
        <button className="btn-action" onClick={addLivrable} style={{ padding: '10px 24px', borderRadius: '30px', fontWeight: 800, display: 'flex', alignItems: 'center', gap: '8px', background: 'var(--main-surface)', border: '1px solid var(--main-border)', color: 'var(--main-text)', cursor: 'pointer' }}>
          <span style={{ fontSize: '1.4rem', lineHeight: '1' }}>+</span> Nouveau Livrable
        </button>
        <button className="btn-Excel" onClick={() => exportToCSV(livrables, 'livrables.csv')} style={{ background: 'var(--excel-color)', color: '#fff', border: 'none', padding: '12px 28px', borderRadius: '30px', fontWeight: 900, cursor: 'pointer', display: 'flex', alignItems: 'center', gap: '10px', boxShadow: '0 8px 16px rgba(70, 155, 111, 0.2)', textTransform: 'uppercase', fontSize: '0.75rem', letterSpacing: '0.05em' }}>
          <span>📥</span> Export Excellence
        </button>
      </div>
      <div style={{ flex: 1, overflow: 'auto', padding: '2.5rem' }}>
        <div style={{ background: 'var(--main-surface)', borderRadius: '24px', border: '1px solid var(--main-border)', overflow: 'hidden', boxShadow: 'var(--card-shadow)' }}>
          <table className="editable-table" style={{ width: '100%', borderCollapse: 'separate', borderSpacing: 0 }}>
            <thead style={{ background: 'var(--excel-color)', color: '#fff' }}>
              <tr style={{ fontWeight: 900, textTransform: 'uppercase', fontSize: '0.75rem', letterSpacing: '0.1em' }}>
                <th style={{ padding: '20px 25px', width: '80px', textAlign: 'center' }}>ID</th>
                <th style={{ padding: '20px 15px', textAlign: 'left' }}>Désignation du livrable / Description</th>
                <th style={{ padding: '20px 15px', textAlign: 'left', width: '250px' }}>Responsabilité Équipe</th>
                <th style={{ padding: '20px 15px', width: '200px', textAlign: 'center' }}>Statut Global</th>
                <th style={{ padding: '20px 15px', width: '140px', textAlign: 'center' }}>Échéance</th>
                <th style={{ padding: '20px 15px', width: '140px', textAlign: 'center' }}>Réalisation</th>
                <th style={{ padding: '20px 15px', width: '70px', textAlign: 'center' }}>Action</th>
              </tr>
            </thead>
            <tbody style={{ color: 'var(--main-text)' }}>
              {livrables.map((l: any, idx: number) => (
                <tr key={l.id} style={{ borderBottom: '1px solid var(--main-border)', background: idx % 2 === 0 ? 'transparent' : 'rgba(0,0,0,0.01)', transition: 'all 0.2s ease' }}>
                  <td style={{ padding: '20px 25px', textAlign: 'center', color: 'var(--text-muted)', fontSize: '0.8rem', fontWeight: 900 }}>{l.id.toString().padStart(2, '0')}</td>
                  <td style={{ padding: '20px 15px' }}>
                    <input value={l.name} onChange={e => handleChange(l.id, 'name', e.target.value)} style={{ width: '100%', border: 'none', background: 'transparent', outline: 'none', fontWeight: 800, color: 'var(--main-text)', fontSize: '1rem', marginBottom: '4px' }} />
                    <div style={{ fontSize: '0.75rem', color: 'var(--text-muted)', fontWeight: 500 }}>{l.type}</div>
                  </td>
                  <td style={{ padding: '20px 15px' }}>
                    <div style={{ display: 'flex', alignItems: 'center', gap: '10px' }}>
                      <div style={{ width: '32px', height: '32px', borderRadius: '10px', background: 'var(--main-bg)', border: '1px solid var(--main-border)', display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: '0.75rem', fontWeight: 900, color: 'var(--primary)' }}>{l.resp.charAt(0)}</div>
                      <input value={l.resp} onChange={e => handleChange(l.id, 'resp', e.target.value)} style={{ flex: 1, border: 'none', background: 'transparent', outline: 'none', color: 'var(--main-text)', fontSize: '0.9rem', fontWeight: 600 }} />
                    </div>
                  </td>
                  <td style={{ padding: '20px 15px', textAlign: 'center' }}>
                    <div style={{
                      background: l.etat === 'Livré' ? 'rgba(16, 185, 129, 0.12)' : l.etat === 'En cours' ? 'rgba(59, 130, 246, 0.12)' : 'rgba(148, 163, 184, 0.12)',
                      color: l.etat === 'Livré' ? '#10b981' : l.etat === 'En cours' ? '#3b82f6' : 'var(--text-muted)',
                      padding: '8px 18px',
                      borderRadius: '30px',
                      fontSize: '0.7rem',
                      fontWeight: 950,
                      display: 'inline-flex',
                      alignItems: 'center',
                      gap: '8px',
                      border: '1px solid currentColor',
                      textTransform: 'uppercase',
                      letterSpacing: '0.05em'
                    }}>
                      <span style={{ width: '8px', height: '8px', borderRadius: '50%', background: 'currentColor' }}></span>
                      <select value={l.etat} onChange={e => handleChange(l.id, 'etat', e.target.value)} style={{ border: 'none', background: 'transparent', color: 'inherit', fontWeight: 'inherit', cursor: 'pointer', outline: 'none' }}>
                        <option value="En attente">En attente</option>
                        <option value="En cours">En cours</option>
                        <option value="Livré">Livré</option>
                      </select>
                    </div>
                  </td>
                  <td style={{ padding: '20px 15px', textAlign: 'center' }}>
                    <input value={l.prev} onChange={e => handleChange(l.id, 'prev', e.target.value)} style={{ width: '100%', border: 'none', background: 'transparent', textAlign: 'center', color: 'var(--main-text)', fontSize: '0.85rem', fontWeight: 700 }} />
                  </td>
                  <td style={{ padding: '20px 15px', textAlign: 'center' }}>
                    <input value={l.real} onChange={e => handleChange(l.id, 'real', e.target.value)} style={{ width: '100%', border: 'none', background: 'transparent', textAlign: 'center', color: l.real !== '-' ? '#10b981' : 'var(--text-muted)', fontSize: '0.85rem', fontWeight: 900 }} />
                  </td>
                  <td style={{ padding: '20px 15px', textAlign: 'center' }}>
                    <button
                      onClick={() => deleteLivrable(l.id)}
                      title="Supprimer ce livrable"
                      style={{ background: 'transparent', border: 'none', cursor: 'pointer', fontSize: '1.1rem', opacity: 0.7, transition: 'opacity 0.2s' }}
                      onMouseEnter={e => e.currentTarget.style.opacity = '1'}
                      onMouseLeave={e => e.currentTarget.style.opacity = '0.7'}
                    >
                      🗑️
                    </button>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
};

const AMDECView = ({ page, projectId, fileId, currentDossierId }: { page: 'echelles' | 'logiciel' | 'materiel' | 'repartition', projectId: string, fileId: string, currentDossierId: string | null }) => {
  const metadata = { projectName: 'FAAB CORE', manager: 'Oliver Martin', version: 'V1.1', lastUpdate: '20/02/2023' };
  const dossierSubKey = currentDossierId ? `-${currentDossierId}` : "";
  const storageKey = `faab-amdec-data-${projectId}-${fileId}${dossierSubKey}`;

  const defaultRows = [
    { id: 1, sys: 'Editer des factures', ssys: 'Domaine facturation', comp: 'Service gestion des ventes', func: 'Charger les données produit et vente', mode: 'Service inaccessible', cause: 'Service externe en maintenance', effet: 'Consolidation impossible', g: 4, f: 3, d: 3, action: 'Monitoring + test service + test d\'intégration' },
    { id: 2, sys: 'Editer des factures', ssys: 'Domaine facturation', comp: 'Repository de données', func: 'Consolider les données', mode: 'Erreur dans la consolidation', cause: 'Bug logiciel', effet: 'Données incorrectes', g: 4, f: 2, d: 3, action: 'Test unitaire' },
    { id: 3, sys: 'Editer des factures', ssys: 'Domaine facturation', comp: 'Service calcul des prix', func: 'Calculer les prix', mode: 'Erreurs de calcul des prix', cause: 'Bug logiciel', effet: 'Erreurs de facturation client', g: 4, f: 2, d: 1, action: 'Test unitaire spécifique' },
    { id: 4, sys: 'Editer des factures', ssys: 'Domaine facturation', comp: 'Templating PDF', func: 'Fusionner les données dans le modèle', mode: 'Mauvais mapping modèle/données', cause: 'Bug logiciel', effet: 'Erreurs de facturation', g: 3, f: 2, d: 2, action: 'Tests d\'intégration' },
    { id: 5, sys: 'Editer des factures', ssys: 'Domaine facturation', comp: 'Templating PDF', func: 'Fusionner les données dans le modèle', mode: 'Volumétrie trop grande pour le modèle', cause: 'Bug logiciel', effet: 'Problème de mise en page', g: 2, f: 3, d: 1, action: 'Test volumétrie + test service' },
    { id: 6, sys: 'Editer des factures', ssys: 'Domaine facturation', comp: 'Service d\'impression PDF', func: 'Éditer le PDF', mode: 'Moteur d\'impression non opérationnel', cause: 'Bug logiciel', effet: 'Pas de sortie PDF', g: 3, f: 2, d: 1, action: 'Test service' }
  ];

  const [rows, setRows] = useState(() => {
    const saved = localStorage.getItem(storageKey);
    return saved ? JSON.parse(saved) : defaultRows;
  });

  useEffect(() => {
    localStorage.setItem(storageKey, JSON.stringify(rows));
  }, [rows, storageKey]);

  if (page === 'echelles') {
    return (
      <div style={{ flex: 1, display: 'flex', flexDirection: 'column', background: 'var(--main-bg)', color: 'var(--main-text)' }}>
        <ExecutionHeader title="26 — Échelles de pondération AMDEC" metadata={metadata} />
        <div style={{ padding: '2.5rem', flex: 1, overflow: 'auto' }}>
          <div style={{ display: 'grid', gridTemplateColumns: 'repeat(2, 1fr)', gap: '2.5rem' }}>
            <div style={{ background: 'var(--main-surface)', padding: '2.5rem', borderRadius: '24px', border: '1px solid var(--main-border)', boxShadow: 'var(--card-shadow)' }}>
              <h3 style={{ marginBottom: '2rem', color: '#ef4444', fontWeight: 900, textTransform: 'uppercase', fontSize: '0.9rem', letterSpacing: '0.1em', display: 'flex', alignItems: 'center', gap: '10px' }}>
                <span style={{ fontSize: '1.2rem' }}>⚖️</span> Gravité (G)
              </h3>
              <table className="editable-table" style={{ width: '100%', borderCollapse: 'separate', borderSpacing: 0 }}>
                <thead>
                  <tr style={{ background: 'rgba(239, 68, 68, 0.1)', color: '#ef4444' }}>
                    <th style={{ padding: '15px 20px', textAlign: 'left', borderRadius: '12px 0 0 12px', fontSize: '0.75rem', fontWeight: 900 }}>NIVEAU</th>
                    <th style={{ padding: '15px 20px', textAlign: 'left', borderRadius: '0 12px 12px 0', fontSize: '0.75rem', fontWeight: 900 }}>DESCRIPTION DE L'IMPACT</th>
                  </tr>
                </thead>
                <tbody>
                  {[
                    { n: 1, d: 'Négligeable (Pas d\'impact perçu)' },
                    { n: 2, d: 'Mineur (Impact faible)' },
                    { n: 3, d: 'Sérieux (Impact significatif)' },
                    { n: 4, d: 'CRITIQUE (Blocage majeur)' }
                  ].map((row, i) => (
                    <tr key={i} style={{ borderBottom: '1px solid var(--main-border)' }}>
                      <td style={{ padding: '15px 20px', fontWeight: 900, color: row.n === 4 ? '#ef4444' : 'inherit' }}>{row.n}</td>
                      <td style={{ padding: '15px 20px', fontWeight: 600, color: row.n === 4 ? '#ef4444' : 'inherit' }}>{row.d}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
            <div style={{ background: 'var(--main-surface)', padding: '2.5rem', borderRadius: '24px', border: '1px solid var(--main-border)', boxShadow: 'var(--card-shadow)' }}>
              <h3 style={{ marginBottom: '2rem', color: '#f59e0b', fontWeight: 900, textTransform: 'uppercase', fontSize: '0.9rem', letterSpacing: '0.1em', display: 'flex', alignItems: 'center', gap: '10px' }}>
                <span style={{ fontSize: '1.2rem' }}>🔄</span> Fréquence (F)
              </h3>
              <table className="editable-table" style={{ width: '100%', borderCollapse: 'separate', borderSpacing: 0 }}>
                <thead>
                  <tr style={{ background: 'rgba(245, 158, 11, 0.1)', color: '#f59e0b' }}>
                    <th style={{ padding: '15px 20px', textAlign: 'left', borderRadius: '12px 0 0 12px', fontSize: '0.75rem', fontWeight: 900 }}>NIVEAU</th>
                    <th style={{ padding: '15px 20px', textAlign: 'left', borderRadius: '0 12px 12px 0', fontSize: '0.75rem', fontWeight: 900 }}>PROBABILITÉ D'OCCURRENCE</th>
                  </tr>
                </thead>
                <tbody>
                  {[
                    { n: 1, d: 'Très Faible (Improbable)' },
                    { n: 2, d: 'Faible (Occasionnel)' },
                    { n: 3, d: 'Modérée (Régulier)' },
                    { n: 4, d: 'ÉLEVÉE (Fréquent)' }
                  ].map((row, i) => (
                    <tr key={i} style={{ borderBottom: '1px solid var(--main-border)' }}>
                      <td style={{ padding: '15px 20px', fontWeight: 900, color: row.n === 4 ? '#f59e0b' : 'inherit' }}>{row.n}</td>
                      <td style={{ padding: '15px 20px', fontWeight: 600, color: row.n === 4 ? '#f59e0b' : 'inherit' }}>{row.d}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>
        </div>
      </div>
    );
  }

  if (page === 'repartition') {
    const critiquesCount = rows.filter((r: any) => (r.g * r.f * r.d) > 20).length;
    const moyensCount = rows.filter((r: any) => { const ip = r.g * r.f * r.d; return ip > 10 && ip <= 20; }).length;
    const faiblesCount = rows.filter((r: any) => (r.g * r.f * r.d) <= 10).length;

    return (
      <div style={{ flex: 1, display: 'flex', flexDirection: 'column', background: 'var(--main-bg)' }}>
        <ExecutionHeader title="26 — Répartition des risques AMDEC" metadata={metadata} />
        <div style={{ padding: '2.5rem' }}>
          <div style={{ background: 'var(--main-surface)', borderRadius: '16px', border: '1px solid var(--main-border)', padding: '2.5rem', boxShadow: 'var(--card-shadow)' }}>
            <h3 style={{ marginBottom: '2.5rem', color: 'var(--main-text)', fontWeight: 800, fontSize: '1.25rem' }}>📊 Synthèse dynamique par niveau de criticité (IPRP = G × F × D)</h3>
            <div style={{ display: 'flex', gap: '2.5rem' }}>
              <div style={{ flex: 1, background: 'rgba(239, 68, 68, 0.08)', padding: '2.5rem', borderRadius: '20px', textAlign: 'center', border: '1px solid rgba(239, 68, 68, 0.2)', transition: 'transform 0.2s', cursor: 'default' }}>
                <div style={{ fontSize: '3.5rem', fontWeight: 900, color: '#ef4444', marginBottom: '10px', textShadow: '0 2px 4px rgba(239, 68, 68, 0.1)' }}>{critiquesCount}</div>
                <div style={{ fontSize: '0.8rem', fontWeight: 800, textTransform: 'uppercase', letterSpacing: '0.1em', color: '#ef4444' }}>🔴 Critiques (IPRP {'>'} 20)</div>
              </div>
              <div style={{ flex: 1, background: 'rgba(245, 158, 11, 0.08)', padding: '2.5rem', borderRadius: '20px', textAlign: 'center', border: '1px solid rgba(245, 158, 11, 0.2)', transition: 'transform 0.2s', cursor: 'default' }}>
                <div style={{ fontSize: '3.5rem', fontWeight: 900, color: '#f59e0b', marginBottom: '10px', textShadow: '0 2px 4px rgba(245, 158, 11, 0.1)' }}>{moyensCount}</div>
                <div style={{ fontSize: '0.8rem', fontWeight: 800, textTransform: 'uppercase', letterSpacing: '0.1em', color: '#f59e0b' }}>🟡 Moyens (10 {'<'} IPRP {"\u2264"} 20)</div>
              </div>
              <div style={{ flex: 1, background: 'rgba(16, 185, 129, 0.08)', padding: '2.5rem', borderRadius: '20px', textAlign: 'center', border: '1px solid rgba(16, 185, 129, 0.2)', transition: 'transform 0.2s', cursor: 'default' }}>
                <div style={{ fontSize: '3.5rem', fontWeight: 900, color: '#10b981', marginBottom: '10px', textShadow: '0 2px 4px rgba(16, 185, 129, 0.1)' }}>{faiblesCount}</div>
                <div style={{ fontSize: '0.8rem', fontWeight: 800, textTransform: 'uppercase', letterSpacing: '0.1em', color: '#10b981' }}>🟢 Faibles (IPRP {"\u2264"} 10)</div>
              </div>
            </div>
          </div>
        </div>
      </div>
    );
  }

  const addAMDEC = () => {
    const nextId = rows.length > 0 ? Math.max(...rows.map((r: any) => r.id || 0)) + 1 : 1;
    const newRow = {
      id: nextId,
      sys: 'Nouveau Système',
      ssys: 'Sous-système',
      comp: 'Composant',
      func: 'Fonction critique',
      mode: 'Défaillance potentielle',
      cause: 'Cause racine',
      effet: 'Impact opérationnel',
      g: 3,
      f: 2,
      d: 2,
      action: 'Action préventive / corrective'
    };
    setRows([...rows, newRow]);
  };

  const deleteAMDEC = (id: number) => {
    setRows(rows.filter((r: any) => r.id !== id));
  };

  const handleChange = (id: number, field: string, val: any) => {
    setRows(rows.map((r: any) => r.id === id ? { ...r, [field]: val } : r));
  };

  return (
    <div style={{ flex: 1, display: 'flex', flexDirection: 'column', background: 'var(--main-bg)', color: 'var(--main-text)' }}>
      <ExecutionHeader title={`26 — AMDEC (${page === 'logiciel' ? 'Software' : 'Hardware'})`} metadata={metadata} />
      <div style={{ flex: 1, padding: '2.5rem', overflow: 'auto' }}>
        <div style={{ display: 'flex', alignItems: 'center', gap: '1.5rem', marginBottom: '2.5rem', justifyContent: 'space-between' }}>
          <div style={{ display: 'flex', gap: '1rem', alignItems: 'center' }}>
            <div style={{ background: 'var(--main-surface)', color: 'var(--main-text)', padding: '14px 28px', borderRadius: '15px', fontWeight: 900, border: '1px solid var(--main-border)', fontSize: '0.85rem', boxShadow: 'var(--card-shadow)', textTransform: 'uppercase', letterSpacing: '0.05em' }}>
              🛡️ Méthode AMDEC Excellence
            </div>
            <div style={{ padding: '14px 28px', background: 'rgba(59, 130, 246, 0.12)', color: '#3b82f6', border: '1px solid rgba(59, 130, 246, 0.2)', borderRadius: '15px', fontSize: '0.85rem', fontWeight: 800 }}>
              Domaine: <strong style={{ color: '#2563eb' }}>{page === 'logiciel' ? 'SOFTWARE / LOGICIEL' : 'HARDWARE / MATÉRIEL'}</strong>
            </div>
          </div>
          <div style={{ display: 'flex', gap: '12px' }}>
            <button
              className="btn-action"
              onClick={addAMDEC}
              style={{ padding: '12px 28px', borderRadius: '35px', fontWeight: 800, display: 'flex', alignItems: 'center', gap: '8px', background: 'var(--main-surface)', border: '1px solid var(--main-border)', color: 'var(--main-text)', cursor: 'pointer', fontSize: '0.8rem', boxShadow: 'var(--card-shadow)' }}
            >
              <span>➕</span> Nouvelle Défaillance
            </button>
            <button
              onClick={() => exportToCSV(rows, `AMDEC_${page}.csv`)}
              style={{ padding: '12px 32px', background: 'var(--excel-color)', color: '#fff', border: 'none', borderRadius: '35px', cursor: 'pointer', fontWeight: 900, fontSize: '0.8rem', display: 'flex', alignItems: 'center', gap: '10px', boxShadow: '0 8px 16px rgba(70, 155, 111, 0.2)', textTransform: 'uppercase', letterSpacing: '0.05em', transition: 'all 0.2s ease' }}
            >
              <span>📥</span> Export Excellence
            </button>
          </div>
        </div>

        <div style={{ background: 'var(--main-surface)', borderRadius: '24px', border: '1px solid var(--main-border)', boxShadow: 'var(--card-shadow)', overflow: 'hidden' }}>
          <table className="editable-table" style={{ width: '100%', borderCollapse: 'separate', borderSpacing: 0, fontSize: '0.9rem' }}>
            <thead style={{ position: 'sticky', top: 0, zIndex: 10 }}>
              <tr style={{ background: 'var(--excel-color)', color: '#fff' }}>
                <th style={{ padding: '20px 15px', fontWeight: 900, textTransform: 'uppercase', letterSpacing: '0.1em', fontSize: '0.7rem' }}>Système / Sous-système</th>
                <th style={{ padding: '20px 15px', fontWeight: 900, textTransform: 'uppercase', letterSpacing: '0.1em', fontSize: '0.7rem' }}>Composants / Fonctions</th>
                <th style={{ padding: '20px 15px', fontWeight: 900, textTransform: 'uppercase', letterSpacing: '0.1em', fontSize: '0.7rem' }}>Mode de Défaillance</th>
                <th style={{ padding: '20px 15px', fontWeight: 900, textTransform: 'uppercase', letterSpacing: '0.1em', fontSize: '0.7rem' }}>Causes & Effets</th>
                <th style={{ padding: '20px 10px', textAlign: 'center', width: '70px', fontWeight: 900 }}>G</th>
                <th style={{ padding: '20px 10px', textAlign: 'center', width: '70px', fontWeight: 900 }}>F</th>
                <th style={{ padding: '20px 10px', textAlign: 'center', width: '70px', fontWeight: 900 }}>D</th>
                <th style={{ padding: '20px 10px', textAlign: 'center', width: '90px', fontWeight: 900 }}>IPRP</th>
                <th style={{ padding: '20px 15px', fontWeight: 900, textTransform: 'uppercase', letterSpacing: '0.1em', fontSize: '0.7rem' }}>Actions Correctives</th>
                <th style={{ padding: '20px 10px', textAlign: 'center', width: '60px', fontWeight: 900 }}>Action</th>
              </tr>
            </thead>
            <tbody style={{ color: 'var(--main-text)' }}>
              {rows.map((r: AMDECRow, i: number) => {
                const iprp = r.g * r.f * r.d;
                return (
                  <tr key={r.id} style={{ borderBottom: '1px solid var(--main-border)', background: i % 2 === 0 ? 'transparent' : 'rgba(0,0,0,0.01)', transition: 'all 0.2s' }}>
                    <td style={{ padding: '15px' }}>
                      <div style={{ fontWeight: 800, color: 'var(--main-text)' }}>{r.sys}</div>
                      <div style={{ fontSize: '0.75rem', color: 'var(--text-muted)', fontWeight: 600 }}>{r.ssys}</div>
                    </td>
                    <td style={{ padding: '15px' }}>
                      <div style={{ fontWeight: 700, color: 'var(--main-text)' }}>{r.comp}</div>
                      <div style={{ fontSize: '0.75rem', color: 'var(--text-muted)', fontStyle: 'italic' }}>{r.func}</div>
                    </td>
                    <td style={{ padding: '15px', fontWeight: 700, color: '#ef4444' }}>{r.mode}</td>
                    <td style={{ padding: '15px' }}>
                      <div style={{ fontSize: '0.8rem', fontWeight: 600 }}><span style={{ color: 'var(--text-muted)' }}>C:</span> {r.cause}</div>
                      <div style={{ fontSize: '0.8rem', fontWeight: 600 }}><span style={{ color: 'var(--text-muted)' }}>E:</span> {r.effet}</div>
                    </td>
                    <td style={{ padding: '15px', textAlign: 'center' }}>
                      <select value={r.g} onChange={e => handleChange(r.id, 'g', parseInt(e.target.value))} style={{ width: '45px', border: '1px solid var(--main-border)', borderRadius: '6px', background: 'var(--main-bg)', color: 'var(--main-text)', padding: '4px', fontWeight: 900 }}>
                        {[1, 2, 3, 4, 5].map(n => <option key={n} value={n}>{n}</option>)}
                      </select>
                    </td>
                    <td style={{ padding: '15px', textAlign: 'center' }}>
                      <select value={r.f} onChange={e => handleChange(r.id, 'f', parseInt(e.target.value))} style={{ width: '45px', border: '1px solid var(--main-border)', borderRadius: '6px', background: 'var(--main-bg)', color: 'var(--main-text)', padding: '4px', fontWeight: 900 }}>
                        {[1, 2, 3, 4, 5].map(n => <option key={n} value={n}>{n}</option>)}
                      </select>
                    </td>
                    <td style={{ padding: '15px', textAlign: 'center' }}>
                      <select value={r.d} onChange={e => handleChange(r.id, 'd', parseInt(e.target.value))} style={{ width: '45px', border: '1px solid var(--main-border)', borderRadius: '6px', background: 'var(--main-bg)', color: 'var(--main-text)', padding: '4px', fontWeight: 900 }}>
                        {[1, 2, 3, 4, 5].map(n => <option key={n} value={n}>{n}</option>)}
                      </select>
                    </td>
                    <td style={{ padding: '15px', textAlign: 'center' }}>
                      <div style={{
                        padding: '6px 12px',
                        background: iprp > 20 ? 'rgba(239, 68, 68, 0.12)' : iprp > 10 ? 'rgba(245, 158, 11, 0.12)' : 'rgba(16, 185, 129, 0.12)',
                        color: iprp > 20 ? '#ef4444' : iprp > 10 ? '#f59e0b' : '#10b981',
                        borderRadius: '8px',
                        fontWeight: 950,
                        fontSize: '0.9rem'
                      }}>
                        {iprp}
                      </div>
                    </td>
                    <td style={{ padding: '15px' }}>
                      <input value={r.action} onChange={e => handleChange(r.id, 'action', e.target.value)} style={{ width: '100%', border: 'none', background: 'transparent', color: '#10b981', fontWeight: 700, fontSize: '0.85rem', outline: 'none' }} />
                    </td>
                    <td style={{ padding: '15px', textAlign: 'center' }}>
                      <button
                        onClick={() => deleteAMDEC(r.id)}
                        title="Supprimer cette défaillance"
                        style={{ background: 'transparent', border: 'none', cursor: 'pointer', fontSize: '1.1rem', opacity: 0.7, transition: 'opacity 0.2s' }}
                        onMouseEnter={e => e.currentTarget.style.opacity = '1'}
                        onMouseLeave={e => e.currentTarget.style.opacity = '0.7'}
                      >
                        🗑️
                      </button>
                    </td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
};

const PlanTestView = ({ page, projectId, fileId, currentDossierId }: { page: 'integration' | 'consolidation', projectId: string, fileId: string, currentDossierId: string | null }) => {
  const metadata = { projectName: 'SAP30 / CMRL', manager: 'Mme Dupont', version: 'V2.3', lastUpdate: '28/11/2022' };
  const dossierSubKey = currentDossierId ? `-${currentDossierId}` : "";
  const storageKey = `faab-tests-data-${projectId}-${fileId}${dossierSubKey}`;

  const [tests, setTests] = useState(() => {
    const saved = localStorage.getItem(storageKey);
    if (saved) return JSON.parse(saved);
    return [
      { id: 1, obj: 'Facturation inter-compagnies', desc: 'Flux de facturation multi-sociétés SAP', cond: 'Commandes cross-company', res: 'Facture générée', stat: 'Conforme', comm: 'Conforme après patch' },
      { id: 2, obj: 'Clôture comptable', desc: 'Validation clôture mensuelle et annuelle', cond: 'Période N-1 clôturée', res: 'Balance validée', stat: 'Conforme', comm: 'Écritures équilibrées' },
      { id: 3, obj: 'Rapprochement des comptes', desc: 'Rapprochement bancaire automatique', cond: 'Extrait bancaire importé', res: 'Lettrage à 98%', stat: 'Conforme', comm: 'RAS' },
      { id: 4, obj: 'Mouvement d\'immobilisation', desc: 'Saisie acquisition et mise au rebut', cond: 'Fiche immo créée', res: 'Tableau d\'amortissement calculé', stat: 'Conforme', comm: 'Règles fiscales OK' },
      { id: 5, obj: 'Amortissements automatiques', desc: 'Lancement batch mensuel d\'amortissement', cond: 'Fin de période', res: 'Dotations passées en compta', stat: 'Conforme', comm: 'Calcul conforme' },
      { id: 6, obj: 'Annulation poste immo', desc: 'Annulation d\'un mouvement erroné', cond: 'Poste non encore clôturé', res: 'Écriture contre-passée', stat: 'Non Conforme', comm: 'Bloqué si lettré' },
      { id: 7, obj: 'Ventes standards', desc: 'Création commande client avec remise', cond: 'Client actif', res: 'Prix net calculé', stat: 'Conforme', comm: 'Barème respecté' },
      { id: 8, obj: 'Unités organisationnelles', desc: 'Affectation division et organisation commerciale', cond: 'Code société CMRL', res: 'Routage correct', stat: 'Conforme', comm: 'OK' },
      { id: 9, obj: 'Prise en commande', desc: 'Saisie EDI commandes clients', cond: 'Flux XML reçu', res: 'Commande intégrée', stat: 'Non Conforme', comm: 'Erreur schéma TVA' },
      { id: 10, obj: 'Facturation client', desc: 'Édition facture PDF et envoi dématérialisé', cond: 'Bon de livraison validé', res: 'Facture PDF générée', stat: 'Conforme', comm: 'Format légal OK' },
      { id: 11, obj: 'Approvisionnement & Achats', desc: 'Création demande d\'achat vers commande fournisseur', cond: 'Seuil réappro atteint', res: 'Bon de commande émis', stat: 'Conforme', comm: 'Circuit signature OK' },
      { id: 12, obj: 'Achats avec contrats-cadres', desc: 'Appel sur contrat de fournitures usine', cond: 'Contrat actif', res: 'Tarif négocié appliqué', stat: 'Non Conforme', comm: 'Plafond dépassé' },
      { id: 13, obj: 'Sous-contractants', desc: 'Gestion de la sous-traitance montage', cond: 'Ordre de fabrication', res: 'Ordre de transfert émis', stat: 'Conforme', comm: 'Flux validé' },
      { id: 14, obj: 'Gestion des expéditions', desc: 'Préparation picking et étiquetage transporteur', cond: 'Ordre de transport', res: 'Bordereau imprimé', stat: 'Conforme', comm: 'Intégration logistique OK' },
      { id: 15, obj: 'Suivi de transport', desc: 'Interface tracking transporteur externe', cond: 'Numéro de colis', res: 'Statut livré actualisé', stat: 'Non Conforme', comm: 'Timeout webhook API' }
    ];
  });

  useEffect(() => {
    localStorage.setItem(storageKey, JSON.stringify(tests));
  }, [tests]);

  const addTest = () => {
    const newId = tests.length > 0 ? Math.max(...tests.map((t: any) => t.id)) + 1 : 1;
    setTests([...tests, { id: newId, obj: 'Nouveau Test', desc: '-', cond: '-', res: '-', stat: 'En attente', comm: '-' }]);
  };

  const deleteTest = (id: number) => {
    setTests(tests.filter((t: any) => t.id !== id));
  };

  const handleChange = (id: number, field: string, val: string) => {
    setTests(tests.map((t: any) => t.id === id ? { ...t, [field]: val } : t));
  };

  if (page === 'consolidation') {
    const conformeCount = tests.filter((t: any) => t.stat === 'Conforme').length;
    const totalCount = tests.length;
    const rate = totalCount > 0 ? Math.round((conformeCount / totalCount) * 100) : 0;

    return (
      <div style={{ flex: 1, display: 'flex', flexDirection: 'column', background: 'var(--main-bg)', color: 'var(--main-text)', fontFamily: "'Outfit', sans-serif" }}>
        <ExecutionHeader title="26 — Consolidation des tests" metadata={metadata} />
        <div style={{ padding: '3.5rem', display: 'grid', gridTemplateColumns: '1fr 1.2fr', gap: '3.5rem', flex: 1 }}>
          <div style={{
            background: 'var(--main-surface)',
            padding: '4rem 3rem',
            borderRadius: '32px',
            border: '1px solid var(--main-border)',
            boxShadow: '0 25px 50px -12px rgba(0, 0, 0, 0.05)',
            textAlign: 'center',
            display: 'flex',
            flexDirection: 'column',
            justifyContent: 'center',
            alignItems: 'center',
            position: 'relative',
            overflow: 'hidden'
          }}>
            <div style={{ position: 'absolute', top: 0, left: 0, width: '100%', height: '8px', background: 'linear-gradient(to right, #10b981, #3b82f6)' }}></div>
            <h3 style={{ fontSize: '0.85rem', color: 'var(--text-muted)', marginBottom: '2.5rem', fontWeight: 950, textTransform: 'uppercase', letterSpacing: '0.2em' }}>INDICE DE QUALITÉ GLOBAL</h3>

            <div style={{ position: 'relative', width: '240px', height: '240px', marginBottom: '2.5rem' }}>
              <svg width="240" height="240" viewBox="0 0 100 100">
                <circle cx="50" cy="50" r="45" fill="none" stroke="var(--main-bg)" strokeWidth="8" />
                <circle cx="50" cy="50" r="45" fill="none" stroke={rate > 80 ? '#10b981' : rate > 50 ? '#f59e0b' : '#ef4444'} strokeWidth="8" strokeDasharray={`${rate * 2.82}, 282`} strokeLinecap="round" transform="rotate(-90 50 50)" style={{ transition: 'stroke-dasharray 1s ease-out' }} />
              </svg>
              <div style={{ position: 'absolute', top: '50%', left: '50%', transform: 'translate(-50%, -50%)', textAlign: 'center' }}>
                <div style={{ fontSize: '4.5rem', fontWeight: 950, color: 'var(--main-text)', letterSpacing: '-2px' }}>{rate}<span style={{ fontSize: '1.5rem', verticalAlign: 'super', opacity: 0.5 }}>%</span></div>
              </div>
            </div>

            <div style={{ color: 'var(--main-text)', fontWeight: 800, fontSize: '1.2rem', background: 'var(--main-bg)', padding: '12px 30px', borderRadius: '20px', border: '1px solid var(--main-border)' }}>
              {conformeCount} Conformes / {totalCount} Tests
            </div>
          </div>

          <div style={{
            background: 'var(--main-surface)',
            padding: '4rem',
            borderRadius: '32px',
            border: '1px solid var(--main-border)',
            boxShadow: '0 25px 50px -12px rgba(0, 0, 0, 0.05)',
            display: 'flex',
            flexDirection: 'column',
            justifyContent: 'center'
          }}>
            <h3 style={{ fontSize: '0.85rem', color: 'var(--text-muted)', marginBottom: '3rem', fontWeight: 950, textTransform: 'uppercase', letterSpacing: '0.2em' }}>RÉPARTITION PAR STATUT</h3>
            <div style={{ display: 'flex', flexDirection: 'column', gap: '35px' }}>
              {[
                { label: 'CONFORME', count: conformeCount, color: '#10b981', icon: '✅' },
                { label: 'NON CONFORME', count: tests.filter((t: any) => t.stat === 'Non Conforme').length, color: '#ef4444', icon: '❌' },
                { label: 'EN ATTENTE', count: tests.filter((t: any) => t.stat === 'En attente').length, color: '#94a3b8', icon: '⏳' }
              ].map(item => (
                <div key={item.label}>
                  <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '12px' }}>
                    <span style={{ color: item.color, fontWeight: 950, fontSize: '0.8rem', display: 'flex', alignItems: 'center', gap: '8px' }}>
                      <span>{item.icon}</span> {item.label}
                    </span>
                    <span style={{ fontWeight: 950, fontSize: '1.2rem' }}>{item.count}</span>
                  </div>
                  <div style={{ height: '14px', background: 'var(--main-bg)', borderRadius: '10px', overflow: 'hidden', border: '1px solid var(--main-border)', padding: '2px' }}>
                    <div style={{
                      width: `${totalCount > 0 ? (item.count / totalCount) * 100 : 0}%`,
                      height: '100%',
                      background: item.color,
                      borderRadius: '8px',
                      transition: 'width 1s cubic-bezier(0.16, 1, 0.3, 1)'
                    }}></div>
                  </div>
                </div>
              ))}
            </div>
          </div>
        </div>
      </div>
    );
  }

  return (
    <div style={{ flex: 1, display: 'flex', flexDirection: 'column', background: 'var(--main-bg)', color: 'var(--main-text)', fontFamily: "'Outfit', sans-serif" }}>
      <ExecutionHeader title="26 — Plan de test (Intégration)" metadata={metadata} />

      <div style={{
        padding: '1.5rem 3rem',
        background: 'rgba(255, 255, 255, 0.02)',
        backdropFilter: 'blur(10px)',
        borderBottom: '1px solid var(--main-border)',
        display: 'flex',
        gap: '20px',
        alignItems: 'center',
        justifyContent: 'space-between'
      }}>
        <div style={{ display: 'flex', alignItems: 'center', gap: '15px' }}>
          <div style={{
            width: '42px',
            height: '42px',
            borderRadius: '12px',
            background: 'linear-gradient(135deg, #3b82f6, #6366f1)',
            display: 'flex',
            alignItems: 'center',
            justifyContent: 'center',
            boxShadow: '0 8px 16px rgba(59, 130, 246, 0.2)'
          }}>
            <span style={{ fontSize: '1.2rem' }}>🧪</span>
          </div>
          <div>
            <h3 style={{ margin: 0, fontSize: '1.1rem', fontWeight: 900, letterSpacing: '-0.01em' }}>PLAN DE TEST INTÉGRATION</h3>
            <p style={{ margin: 0, fontSize: '0.75rem', color: 'var(--text-muted)', fontWeight: 600 }}>Validation fonctionnelle & technique</p>
          </div>
        </div>

        <div style={{ display: 'flex', gap: '12px' }}>
          <button className="btn-action" onClick={addTest} style={{
            padding: '10px 24px',
            borderRadius: '12px',
            fontWeight: 800,
            display: 'flex',
            alignItems: 'center',
            gap: '8px',
            background: 'var(--main-surface)',
            border: '1px solid var(--main-border)',
            color: 'var(--main-text)',
            cursor: 'pointer',
            fontSize: '0.8rem'
          }}>
            <span>➕</span> Nouveau Test
          </button>
          <button className="btn-Excel" onClick={() => exportToCSV(tests, 'plan_tests.csv')} style={{
            padding: '10px 24px',
            background: 'var(--excel-color)',
            color: '#fff',
            border: 'none',
            borderRadius: '12px',
            fontWeight: 900,
            cursor: 'pointer',
            display: 'flex',
            alignItems: 'center',
            gap: '10px',
            boxShadow: '0 8px 16px rgba(70, 155, 111, 0.2)',
            textTransform: 'uppercase',
            fontSize: '0.8rem'
          }}>
            <span>📥</span> EXPORT SMART
          </button>
        </div>
      </div>

      <div style={{ flex: 1, overflow: 'auto', padding: '2.5rem 3rem' }}>
        <div style={{
          background: 'var(--main-surface)',
          borderRadius: '24px',
          border: '1px solid var(--main-border)',
          boxShadow: '0 20px 40px rgba(0,0,0,0.05)',
          overflow: 'hidden'
        }}>
          <table style={{ width: '100%', borderCollapse: 'separate', borderSpacing: 0 }}>
            <thead>
              <tr style={{ background: 'rgba(59, 130, 246, 0.03)' }}>
                <th style={{ padding: '24px 25px', color: 'var(--text-muted)', fontWeight: 950, textTransform: 'uppercase', letterSpacing: '0.15em', fontSize: '0.7rem', width: '300px' }}>Objet du Test</th>
                <th style={{ padding: '24px 15px', color: 'var(--text-muted)', fontWeight: 950, textTransform: 'uppercase', letterSpacing: '0.15em', fontSize: '0.7rem' }}>Description & Conditions</th>
                <th style={{ padding: '24px 15px', color: 'var(--text-muted)', fontWeight: 950, textTransform: 'uppercase', letterSpacing: '0.15em', fontSize: '0.7rem' }}>Résultat Obtenu</th>
                <th style={{ padding: '24px 25px', color: 'var(--text-muted)', fontWeight: 950, textTransform: 'uppercase', letterSpacing: '0.15em', fontSize: '0.7rem', textAlign: 'center', width: '220px' }}>État de Conformité</th>
                <th style={{ padding: '24px 15px', color: 'var(--text-muted)', fontWeight: 950, textTransform: 'uppercase', letterSpacing: '0.15em', fontSize: '0.7rem', textAlign: 'center', width: '70px' }}>Action</th>
              </tr>
            </thead>
            <tbody style={{ color: 'var(--main-text)' }}>
              {tests.map((t: any, i: number) => (
                <tr key={t.id} className="test-row" style={{
                  borderTop: '1px solid var(--main-border)',
                  background: 'transparent',
                  transition: 'all 0.2s ease'
                }}>
                  <td style={{ padding: '25px' }}>
                    <input value={t.obj} onChange={e => handleChange(t.id, 'obj', e.target.value)} style={{ width: '100%', border: 'none', background: 'transparent', fontWeight: 800, color: 'var(--main-text)', fontSize: '1rem', outline: 'none' }} />
                  </td>
                  <td style={{ padding: '25px 15px' }}>
                    <textarea value={t.desc} onChange={e => handleChange(t.id, 'desc', e.target.value)} style={{ width: '100%', border: 'none', background: 'transparent', color: 'var(--main-text)', fontSize: '0.85rem', fontWeight: 600, marginBottom: '6px', resize: 'none', outline: 'none' }} rows={2} />
                    <div style={{ fontSize: '0.7rem', color: 'var(--text-muted)', fontWeight: 700, display: 'flex', alignItems: 'center', gap: '6px' }}>
                      <span style={{ color: 'var(--primary)' }}>PRE-COND :</span> {t.cond}
                    </div>
                  </td>
                  <td style={{ padding: '25px 15px' }}>
                    <textarea value={t.res} onChange={e => handleChange(t.id, 'res', e.target.value)} style={{ width: '100%', border: 'none', background: 'transparent', color: 'var(--main-text)', fontSize: '0.85rem', fontWeight: 700, outline: 'none', resize: 'none' }} rows={2} />
                  </td>
                  <td style={{ padding: '25px', textAlign: 'center' }}>
                    <div style={{
                      padding: '4px 8px',
                      borderRadius: '16px',
                      background: t.stat === 'Conforme' ? 'rgba(16, 185, 129, 0.1)' : t.stat === 'Non Conforme' ? 'rgba(239, 68, 68, 0.1)' : 'rgba(148, 163, 184, 0.1)',
                      color: t.stat === 'Conforme' ? '#10b981' : t.stat === 'Non Conforme' ? '#ef4444' : 'var(--text-muted)',
                      border: '1px solid currentColor',
                      display: 'inline-block',
                      transition: 'all 0.3s'
                    }}>
                      <select
                        value={t.stat}
                        onChange={e => handleChange(t.id, 'stat', e.target.value)}
                        style={{ border: 'none', background: 'transparent', color: 'inherit', fontWeight: 950, cursor: 'pointer', outline: 'none', textAlignLast: 'center', fontSize: '0.75rem', padding: '6px 15px', textTransform: 'uppercase', letterSpacing: '0.05em' }}
                      >
                        <option value="Conforme">Conforme</option>
                        <option value="Non Conforme">Non Conforme</option>
                        <option value="En attente">En attente</option>
                      </select>
                    </div>
                  </td>
                  <td style={{ padding: '25px 15px', textAlign: 'center' }}>
                    <button
                      onClick={() => deleteTest(t.id)}
                      title="Supprimer ce test"
                      style={{ background: 'transparent', border: 'none', cursor: 'pointer', fontSize: '1.1rem', opacity: 0.7, transition: 'opacity 0.2s' }}
                      onMouseEnter={e => e.currentTarget.style.opacity = '1'}
                      onMouseLeave={e => e.currentTarget.style.opacity = '0.7'}
                    >
                      🗑️
                    </button>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>
      <style>{`
        .test-row:hover {
          background: rgba(59, 130, 246, 0.02) !important;
        }
        .test-row:hover input, .test-row:hover textarea {
          color: var(--primary) !important;
        }
      `}</style>
    </div>
  );
};

const ChangementView = ({ page, projectId, fileId, currentDossierId }: { page: 'demandes' | 'parametrage', projectId: string, fileId: string, currentDossierId: string | null }) => {
  const metadata = { projectName: 'SAP30 / CMRL', manager: 'Mme Dupont', version: 'V2.3', lastUpdate: '25/7/2023' };
  const dossierSubKey = currentDossierId ? `-${currentDossierId}` : "";
  const storageKey = `faab-chg-data-${projectId}-${fileId}${dossierSubKey}`;

  const defaultRows = [
    { id: 1, date: '06/04/2023', demandeur: 'Service clients', desc: 'Inclure un changement du matériel informatique dans le budget', raison: 'Amélioration des postes de travail', impact: 'Périmètre', prio: 'P1', status: 'En attente' },
    { id: 2, date: '08/04/2023', demandeur: 'Direction Ventes', desc: 'Inclure CRM dans le périmètre du projet', raison: 'Efficacité des vendeurs', impact: 'Échéancier', prio: 'P1', status: 'Refusé' },
    { id: 3, date: '15/05/2023', demandeur: 'Production', desc: 'Inclure les chefs d\'équipe dans les utilisateurs administratifs', raison: 'Besoins métiers', impact: 'Communication', prio: 'P2', status: 'Validé' },
  ];

  const [rows, setRows] = useState(() => {
    const saved = localStorage.getItem(storageKey);
    return saved ? JSON.parse(saved) : defaultRows;
  });

  useEffect(() => {
    localStorage.setItem(storageKey, JSON.stringify(rows));
  }, [rows, storageKey]);

  const addChangement = () => {
    const nextId = rows.length > 0 ? Math.max(...rows.map((r: any) => Number(r.id) || 0)) + 1 : 1;
    const today = new Date().toLocaleDateString('fr-FR');
    const newRow = {
      id: nextId,
      date: today,
      demandeur: 'Équipe Projet',
      desc: 'Nouvelle demande de modification de périmètre',
      raison: 'Optimisation opérationnelle',
      impact: 'Périmètre',
      prio: 'P2',
      status: 'En attente'
    };
    setRows([...rows, newRow]);
  };

  const deleteChangement = (id: number | string) => {
    setRows(rows.filter((r: any) => r.id !== id));
  };

  const handleChange = (id: number | string, field: string, val: any) => {
    setRows(rows.map((r: any) => r.id === id ? { ...r, [field]: val } : r));
  };

  if (page === 'parametrage') {
    return (
      <div style={{ flex: 1, display: 'flex', flexDirection: 'column', background: 'var(--main-bg)', color: 'var(--main-text)', fontFamily: "'Outfit', sans-serif" }}>
        <ExecutionHeader title="25 — Paramétrage du Registre" metadata={metadata} />
        <div style={{ flex: 1, padding: '2.5rem 3rem', overflow: 'auto' }}>
          <div style={{ background: 'var(--main-surface)', borderRadius: '24px', border: '1px solid var(--main-border)', padding: '3.5rem', boxShadow: '0 20px 40px rgba(0,0,0,0.05)' }}>
            <div style={{ marginBottom: '4rem', display: 'flex', alignItems: 'center', gap: '20px' }}>
              <div style={{
                width: '56px',
                height: '56px',
                borderRadius: '16px',
                background: 'linear-gradient(135deg, #10b981, #3b82f6)',
                display: 'flex',
                alignItems: 'center',
                justifyContent: 'center',
                boxShadow: '0 8px 16px rgba(16, 185, 129, 0.2)'
              }}>
                <span style={{ fontSize: '1.8rem' }}>⚙️</span>
              </div>
              <div>
                <h3 style={{ margin: 0, fontSize: '1.5rem', fontWeight: 950, textTransform: 'uppercase', color: 'var(--excel-color)', letterSpacing: '-0.02em' }}>Configuration des Changements</h3>
                <p style={{ margin: 0, fontSize: '0.9rem', color: 'var(--text-muted)', fontWeight: 600 }}>Définition des règles et processus de gestion</p>
              </div>
            </div>

            <div style={{ display: 'grid', gridTemplateColumns: '1fr 1.2fr', gap: '4rem' }}>
              <section>
                <h4 style={{ fontSize: '0.8rem', color: 'var(--text-muted)', marginBottom: '2rem', fontWeight: 950, textTransform: 'uppercase', letterSpacing: '0.15em', display: 'flex', alignItems: 'center', gap: '10px' }}>
                  <div style={{ width: '20px', height: '2px', background: 'var(--excel-color)' }}></div>
                  TYPES D'IMPACT ANALYSÉS
                </h4>
                <div style={{ display: 'flex', flexDirection: 'column', gap: '12px' }}>
                  {[
                    { label: 'Périmètre fonctionnel', icon: '🎯' },
                    { label: 'Échéancier critique', icon: '📅' },
                    { label: 'Budget & Coûts', icon: '💰' },
                    { label: 'Qualité livrables', icon: '✨' },
                    { label: 'Ressources humaines', icon: '👥' }
                  ].map(impact => (
                    <div key={impact.label} style={{
                      padding: '16px 24px',
                      background: 'var(--main-bg)',
                      borderRadius: '16px',
                      fontSize: '0.95rem',
                      fontWeight: 800,
                      border: '1px solid var(--main-border)',
                      display: 'flex',
                      alignItems: 'center',
                      gap: '15px',
                      transition: 'all 0.2s ease',
                      cursor: 'default'
                    }}>
                      <span style={{ fontSize: '1.2rem' }}>{impact.icon}</span>
                      {impact.label}
                    </div>
                  ))}
                </div>
              </section>

              <section>
                <h4 style={{ fontSize: '0.8rem', color: 'var(--text-muted)', marginBottom: '2rem', fontWeight: 950, textTransform: 'uppercase', letterSpacing: '0.15em', display: 'flex', alignItems: 'center', gap: '10px' }}>
                  <div style={{ width: '20px', height: '2px', background: '#3b82f6' }}></div>
                  WORKFLOW DE VALIDATION
                </h4>
                <div style={{
                  padding: '35px',
                  border: '1px dashed var(--main-border)',
                  borderRadius: '24px',
                  background: 'rgba(59, 130, 246, 0.02)',
                  display: 'flex',
                  flexDirection: 'column',
                  gap: '5px'
                }}>
                  {[
                    { step: 1, title: 'Soumission initiale', actor: 'Demandeur', color: '#10b981' },
                    { step: 2, title: "Analyse d'impact", actor: 'Chef de projet', color: '#3b82f6' },
                    { step: 3, title: 'Arbitrage Décisionnel', actor: 'Comité (COPIL)', color: '#6366f1' },
                    { step: 4, title: 'Mise en œuvre', actor: 'Équipe Projet', color: '#8b5cf6' }
                  ].map((w, i, arr) => (
                    <Fragment key={w.step}>
                      <div style={{ display: 'flex', alignItems: 'center', gap: '20px' }}>
                        <div style={{
                          width: '40px',
                          height: '40px',
                          borderRadius: '50%',
                          background: w.color,
                          color: '#fff',
                          display: 'flex',
                          alignItems: 'center',
                          justifyContent: 'center',
                          fontWeight: 950,
                          boxShadow: `0 8px 16px ${w.color}33`,
                          fontSize: '0.9rem'
                        }}>{w.step}</div>
                        <div>
                          <div style={{ fontWeight: 900, fontSize: '1rem', color: 'var(--main-text)' }}>{w.title}</div>
                          <div style={{ fontSize: '0.75rem', color: 'var(--text-muted)', fontWeight: 700 }}>Acteur : {w.actor}</div>
                        </div>
                      </div>
                      {i < arr.length - 1 && (
                        <div style={{ height: '30px', width: '2px', background: 'linear-gradient(to bottom, ' + w.color + ', ' + arr[i + 1].color + ')', marginLeft: '19px', margin: '5px 0' }}></div>
                      )}
                    </Fragment>
                  ))}
                </div>
              </section>
            </div>
          </div>
        </div>
      </div>
    );
  }

  return (
    <div style={{ flex: 1, display: 'flex', flexDirection: 'column', background: 'var(--main-bg)', color: 'var(--main-text)', fontFamily: "'Outfit', sans-serif" }}>
      <ExecutionHeader title="25 — Registre demandes de changement" metadata={metadata} />

      <div style={{
        padding: '1.5rem 3rem',
        background: 'rgba(255, 255, 255, 0.02)',
        backdropFilter: 'blur(10px)',
        borderBottom: '1px solid var(--main-border)',
        display: 'flex',
        gap: '20px',
        alignItems: 'center',
        justifyContent: 'space-between'
      }}>
        <div style={{ display: 'flex', alignItems: 'center', gap: '15px' }}>
          <div style={{
            width: '42px',
            height: '42px',
            borderRadius: '12px',
            background: 'linear-gradient(135deg, #10b981, #059669)',
            display: 'flex',
            alignItems: 'center',
            justifyContent: 'center',
            boxShadow: '0 8px 16px rgba(16, 185, 129, 0.2)'
          }}>
            <span style={{ fontSize: '1.2rem' }}>🔄</span>
          </div>
          <div>
            <h3 style={{ margin: 0, fontSize: '1.1rem', fontWeight: 900, letterSpacing: '-0.01em' }}>REGISTRE DES CHANGEMENTS</h3>
            <p style={{ margin: 0, fontSize: '0.75rem', color: 'var(--text-muted)', fontWeight: 600 }}>Suivi des modifications de périmètre</p>
          </div>
        </div>

        <div style={{ display: 'flex', gap: '12px' }}>
          <button
            className="btn-action"
            onClick={addChangement}
            style={{
              padding: '10px 24px',
              borderRadius: '12px',
              fontWeight: 800,
              display: 'flex',
              alignItems: 'center',
              gap: '8px',
              background: 'var(--main-surface)',
              border: '1px solid var(--main-border)',
              color: 'var(--main-text)',
              cursor: 'pointer',
              fontSize: '0.8rem'
            }}
          >
            <span>➕</span> Nouveau Changement
          </button>
          <button className="btn-Excel" onClick={() => exportToCSV(rows, 'registre_changements.csv')} style={{
            padding: '10px 24px',
            background: 'var(--excel-color)',
            color: '#fff',
            border: 'none',
            borderRadius: '12px',
            fontWeight: 900,
            cursor: 'pointer',
            display: 'flex',
            alignItems: 'center',
            gap: '10px',
            boxShadow: '0 8px 16px rgba(16, 185, 129, 0.2)',
            textTransform: 'uppercase',
            fontSize: '0.8rem'
          }}>
            <span>📥</span> EXPORT EXPERT
          </button>
        </div>
      </div>

      <div style={{ flex: 1, padding: '2.5rem 3rem', overflow: 'auto' }}>
        <div style={{
          background: 'var(--main-surface)',
          borderRadius: '24px',
          border: '1px solid var(--main-border)',
          boxShadow: '0 20px 40px rgba(0,0,0,0.05)',
          overflow: 'hidden'
        }}>
          <table style={{ width: '100%', borderCollapse: 'separate', borderSpacing: 0 }}>
            <thead>
              <tr style={{ background: 'rgba(59, 130, 246, 0.03)' }}>
                <th style={{ padding: '24px 25px', width: '80px', textAlign: 'center', fontWeight: 950, fontSize: '0.7rem', textTransform: 'uppercase', color: 'var(--text-muted)', letterSpacing: '0.15em' }}>ID</th>
                <th style={{ padding: '24px 15px', textAlign: 'center', width: '120px', fontWeight: 950, fontSize: '0.7rem', textTransform: 'uppercase', color: 'var(--text-muted)', letterSpacing: '0.15em' }}>Date</th>
                <th style={{ padding: '24px 15px', textAlign: 'left', fontWeight: 950, fontSize: '0.7rem', textTransform: 'uppercase', color: 'var(--text-muted)', letterSpacing: '0.15em' }}>Demandeur</th>
                <th style={{ padding: '24px 15px', textAlign: 'left', fontWeight: 950, fontSize: '0.7rem', textTransform: 'uppercase', color: 'var(--text-muted)', letterSpacing: '0.15em' }}>Description & Motivations</th>
                <th style={{ padding: '24px 15px', textAlign: 'center', fontWeight: 950, fontSize: '0.7rem', textTransform: 'uppercase', color: 'var(--text-muted)', letterSpacing: '0.15em', width: '150px' }}>Nature Impact</th>
                <th style={{ padding: '24px 30px', fontWeight: 950, fontSize: '0.7rem', textTransform: 'uppercase', color: 'var(--text-muted)', letterSpacing: '0.15em', textAlign: 'center', width: '200px' }}>Décision</th>
                <th style={{ padding: '24px 15px', color: 'var(--text-muted)', fontWeight: 950, textTransform: 'uppercase', letterSpacing: '0.15em', fontSize: '0.7rem', textAlign: 'center', width: '70px' }}>Action</th>
              </tr>
            </thead>
            <tbody style={{ color: 'var(--main-text)' }}>
              {rows.map((c: any) => (
                <tr key={c.id} className="change-row" style={{
                  borderTop: '1px solid var(--main-border)',
                  background: 'transparent',
                  transition: 'all 0.2s ease'
                }}>
                  <td style={{ padding: '25px', textAlign: 'center', fontWeight: 950, color: 'var(--text-muted)', fontSize: '0.9rem' }}>{c.id.toString().padStart(2, '0')}</td>
                  <td style={{ padding: '25px 15px', textAlign: 'center' }}>
                    <input
                      value={c.date}
                      onChange={e => handleChange(c.id, 'date', e.target.value)}
                      style={{ padding: '6px 12px', background: 'var(--main-bg)', borderRadius: '8px', border: '1px solid var(--main-border)', fontWeight: 800, fontSize: '0.85rem', color: 'var(--main-text)', width: '100px', textAlign: 'center' }}
                    />
                  </td>
                  <td style={{ padding: '25px 15px' }}>
                    <div style={{ display: 'flex', alignItems: 'center', gap: '10px' }}>
                      <div style={{ width: '32px', height: '32px', borderRadius: '50%', background: 'linear-gradient(135deg, #3b82f6, #6366f1)', color: '#fff', display: 'flex', alignItems: 'center', justifyContent: 'center', fontWeight: 900, fontSize: '0.75rem', flexShrink: 0 }}>
                        {c.demandeur?.charAt(0) || 'D'}
                      </div>
                      <input
                        value={c.demandeur}
                        onChange={e => handleChange(c.id, 'demandeur', e.target.value)}
                        style={{ fontWeight: 800, fontSize: '0.95rem', border: 'none', background: 'transparent', color: 'var(--main-text)', outline: 'none', width: '100%' }}
                      />
                    </div>
                  </td>
                  <td style={{ padding: '25px 15px' }}>
                    <input
                      value={c.desc}
                      onChange={e => handleChange(c.id, 'desc', e.target.value)}
                      style={{ fontWeight: 850, marginBottom: '6px', fontSize: '1rem', color: 'var(--main-text)', border: 'none', background: 'transparent', outline: 'none', width: '100%' }}
                    />
                    <div style={{ display: 'flex', alignItems: 'center', gap: '6px' }}>
                      <span style={{ fontSize: '0.8rem', color: 'var(--text-muted)', fontStyle: 'italic', fontWeight: 600 }}>Motif :</span>
                      <input
                        value={c.raison}
                        onChange={e => handleChange(c.id, 'raison', e.target.value)}
                        style={{ fontSize: '0.8rem', color: 'var(--text-muted)', fontStyle: 'italic', fontWeight: 600, border: 'none', background: 'transparent', outline: 'none', width: '100%' }}
                      />
                    </div>
                  </td>
                  <td style={{ padding: '25px 15px', textAlign: 'center' }}>
                    <select
                      value={c.impact}
                      onChange={e => handleChange(c.id, 'impact', e.target.value)}
                      style={{ padding: '6px 14px', background: 'rgba(59, 130, 246, 0.05)', color: '#3b82f6', borderRadius: '8px', fontSize: '0.75rem', fontWeight: 900, textTransform: 'uppercase', border: '1px solid rgba(59, 130, 246, 0.2)', cursor: 'pointer', outline: 'none' }}
                    >
                      <option value="Périmètre">Périmètre</option>
                      <option value="Échéancier">Échéancier</option>
                      <option value="Budget">Budget</option>
                      <option value="Qualité">Qualité</option>
                      <option value="Communication">Communication</option>
                    </select>
                  </td>
                  <td style={{ padding: '25px 30px', textAlign: 'center' }}>
                    <div style={{
                      padding: '8px 18px',
                      borderRadius: '30px',
                      fontSize: '0.75rem',
                      fontWeight: 950,
                      background: c.status === 'Validé' || c.status === 'validée' ? 'linear-gradient(135deg, #10b981, #059669)' : c.status === 'Refusé' ? 'linear-gradient(135deg, #ef4444, #dc2626)' : 'linear-gradient(135deg, #f59e0b, #d97706)',
                      color: '#fff',
                      boxShadow: '0 4px 12px rgba(0,0,0,0.1)',
                      textTransform: 'uppercase',
                      letterSpacing: '0.05em',
                      display: 'inline-block'
                    }}>
                      <select
                        value={c.status === 'validée' ? 'Validé' : c.status}
                        onChange={e => handleChange(c.id, 'status', e.target.value)}
                        style={{ border: 'none', background: 'transparent', color: '#fff', fontWeight: 950, cursor: 'pointer', outline: 'none', textTransform: 'uppercase' }}
                      >
                        <option value="En attente" style={{ background: 'var(--main-surface)', color: 'var(--main-text)' }}>En attente</option>
                        <option value="Validé" style={{ background: 'var(--main-surface)', color: 'var(--main-text)' }}>Validé</option>
                        <option value="Refusé" style={{ background: 'var(--main-surface)', color: 'var(--main-text)' }}>Refusé</option>
                      </select>
                    </div>
                  </td>
                  <td style={{ padding: '25px 15px', textAlign: 'center' }}>
                    <button
                      onClick={() => deleteChangement(c.id)}
                      title="Supprimer ce changement"
                      style={{ background: 'transparent', border: 'none', cursor: 'pointer', fontSize: '1.1rem', opacity: 0.7, transition: 'opacity 0.2s' }}
                      onMouseEnter={e => e.currentTarget.style.opacity = '1'}
                      onMouseLeave={e => e.currentTarget.style.opacity = '0.7'}
                    >
                      🗑️
                    </button>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>
      <style>{`
        .change-row:hover {
          background: rgba(16, 185, 129, 0.02) !important;
          transform: translateY(-1px);
        }
      `}</style>
    </div>
  );
};

const ActionsView = ({ page, projectId, fileId, currentDossierId }: { page: 'actions' | 'parametrage', projectId: string, fileId: string, currentDossierId: string | null }) => {
  const metadata = { projectName: 'SAP30 / CMRL', manager: 'Mme Dupont', version: 'V2.3', lastUpdate: '28/11/2022' };
  const dossierSubKey = currentDossierId ? `-${currentDossierId}` : "";
  const storageKey = `faab-actions-data-${projectId}-${fileId}${dossierSubKey}`;

  const [actions, setActions] = useState(() => {
    const saved = localStorage.getItem(storageKey);
    if (saved) return JSON.parse(saved);
    return [
      { id: 1, action: 'Organiser réunion de lancement', origin: 'Charte', resp: 'Chantal Dupont', ech: '15/07/2022', real: '100%', prio: 'P1', etat: 'Réalisé' },
      { id: 2, action: 'Valider le cahier des charges avec le sponsor', origin: 'Charte', resp: 'Olivier Martin', ech: '30/07/2022', real: '100%', prio: 'P1', etat: 'Réalisé' },
      { id: 3, action: 'Finaliser AMDEC logiciel', origin: 'Gestion opérationnelle', resp: 'Eric Duteil', ech: '30/08/2022', real: '75%', prio: 'P1', etat: 'En cours' },
      { id: 4, action: 'Organiser session de formation utilisateurs clés', origin: 'Instances COPIL', resp: 'Monique Davant', ech: '15/09/2022', real: '40%', prio: 'P2', etat: 'En cours' },
      { id: 5, action: 'Exécuter tests d\'intégration module Finance', origin: 'Audit Qualité', resp: 'Denise Attali', ech: '30/09/2022', real: '60%', prio: 'P1', etat: 'En cours' },
      { id: 6, action: 'Exécuter tests d\'intégration module Ventes', origin: 'Audit Qualité', resp: 'Jean Albert', ech: '31/10/2022', real: '0%', prio: 'P1', etat: 'À faire' },
      { id: 7, action: 'Backup complet de la base de données prod', origin: 'Réunions Techniques', resp: 'Daniel Tech', ech: '15/10/2022', real: '0%', prio: 'P1', etat: 'À faire' },
      { id: 8, action: 'Rédiger rapport de suivi v2', origin: 'Instances COPIL', resp: 'Chantal Dupont', ech: '31/10/2022', real: '0%', prio: 'P2', etat: 'À faire' }
    ];
  });

  useEffect(() => {
    localStorage.setItem(storageKey, JSON.stringify(actions));
  }, [actions]);

  const addAction = () => {
    const newId = actions.length > 0 ? Math.max(...actions.map((a: any) => a.id)) + 1 : 1;
    setActions([...actions, { id: newId, action: 'Nouvelle Action', origin: 'Gestion opérationnelle', resp: 'Responsable', ech: '31/12/2022', real: '0%', prio: 'P2', etat: 'À faire' }]);
  };

  const deleteAction = (id: number) => {
    setActions(actions.filter((a: any) => a.id !== id));
  };

  const handleChange = (id: number, field: string, val: string) => {
    setActions(actions.map((a: any) => a.id === id ? { ...a, [field]: val } : a));
  };

  const totalActions = actions.length;
  const termineesCount = actions.filter((a: any) => a.etat === 'Réalisé' || a.etat === 'Terminé').length;
  const enCoursCount = actions.filter((a: any) => a.etat === 'En cours').length;
  const aFaireCount = actions.filter((a: any) => a.etat === 'À faire' || a.etat === 'En attente').length;

  if (page === 'parametrage') {
    return (
      <div style={{ flex: 1, display: 'flex', flexDirection: 'column', background: 'var(--main-bg)', color: 'var(--main-text)' }}>
        <ExecutionHeader title="22 — Paramétrage des actions" metadata={metadata} />
        <div style={{ flex: 1, padding: '2.5rem', overflow: 'auto' }}>
          <div style={{ background: 'var(--main-surface)', borderRadius: '24px', border: '1px solid var(--main-border)', padding: '3rem', boxShadow: 'var(--card-shadow)' }}>
            <h3 style={{ color: 'var(--excel-color)', marginBottom: '3rem', display: 'flex', alignItems: 'center', gap: '15px', fontSize: '1.4rem', fontWeight: 950, textTransform: 'uppercase' }}>
              <span style={{ fontSize: '1.8rem' }}>⚙️</span> Configuration des Sources
            </h3>
            <div style={{ display: 'grid', gridTemplateColumns: '1fr', gap: '3rem', maxWidth: '600px' }}>
              <section>
                <h4 style={{ fontSize: '0.8rem', color: 'var(--text-muted)', marginBottom: '1.5rem', fontWeight: 900, textTransform: 'uppercase', letterSpacing: '0.1em' }}>ORIGINES POSSIBLES</h4>
                <div style={{ display: 'flex', flexWrap: 'wrap', gap: '1rem' }}>
                  {['Charte Projet', 'Gestion opérationnelle', 'Instances COPIL', 'Audit Qualité', 'Réunions Techniques'].map(o => (
                    <div key={o} style={{ padding: '14px 24px', background: 'var(--main-bg)', borderRadius: '12px', fontSize: '0.9rem', fontWeight: 700, border: '1px solid var(--main-border)' }}>{o}</div>
                  ))}
                </div>
              </section>
            </div>
          </div>
        </div>
      </div>
    );
  }

  return (
    <div style={{ flex: 1, display: 'flex', flexDirection: 'column', background: 'var(--main-bg)', color: 'var(--main-text)' }}>
      <ExecutionHeader title="22 — Suivi des actions" metadata={metadata} />

      {/* Dynamic KPI Row */}
      <div style={{ padding: '1.5rem 2.5rem 0', display: 'grid', gridTemplateColumns: 'repeat(4, 1fr)', gap: '1.25rem' }}>
        <div style={{ background: 'var(--main-surface)', padding: '1.25rem', borderRadius: '14px', border: '1px solid var(--main-border)', borderLeft: '4px solid #3b82f6', boxShadow: 'var(--card-shadow)' }}>
          <div style={{ fontSize: '0.75rem', fontWeight: 800, color: 'var(--text-muted)', textTransform: 'uppercase' }}>Total Actions</div>
          <div style={{ fontSize: '1.8rem', fontWeight: 950, color: '#3b82f6', marginTop: '4px' }}>{totalActions}</div>
          <div style={{ fontSize: '0.75rem', color: 'var(--text-muted)' }}>Journal consolidé</div>
        </div>
        <div style={{ background: 'var(--main-surface)', padding: '1.25rem', borderRadius: '14px', border: '1px solid var(--main-border)', borderLeft: '4px solid #10b981', boxShadow: 'var(--card-shadow)' }}>
          <div style={{ fontSize: '0.75rem', fontWeight: 800, color: 'var(--text-muted)', textTransform: 'uppercase' }}>Terminées ✅</div>
          <div style={{ fontSize: '1.8rem', fontWeight: 950, color: '#10b981', marginTop: '4px' }}>{termineesCount}</div>
          <div style={{ fontSize: '0.75rem', color: '#10b981', fontWeight: 700 }}>Actions achevées</div>
        </div>
        <div style={{ background: 'var(--main-surface)', padding: '1.25rem', borderRadius: '14px', border: '1px solid var(--main-border)', borderLeft: '4px solid #3b82f6', boxShadow: 'var(--card-shadow)' }}>
          <div style={{ fontSize: '0.75rem', fontWeight: 800, color: 'var(--text-muted)', textTransform: 'uppercase' }}>En Cours 🔵</div>
          <div style={{ fontSize: '1.8rem', fontWeight: 950, color: '#3b82f6', marginTop: '4px' }}>{enCoursCount}</div>
          <div style={{ fontSize: '0.75rem', color: '#3b82f6', fontWeight: 700 }}>Travail en cours</div>
        </div>
        <div style={{ background: 'var(--main-surface)', padding: '1.25rem', borderRadius: '14px', border: '1px solid var(--main-border)', borderLeft: '4px solid #f59e0b', boxShadow: 'var(--card-shadow)' }}>
          <div style={{ fontSize: '0.75rem', fontWeight: 800, color: 'var(--text-muted)', textTransform: 'uppercase' }}>À Faire ⬜</div>
          <div style={{ fontSize: '1.8rem', fontWeight: 950, color: '#f59e0b', marginTop: '4px' }}>{aFaireCount}</div>
          <div style={{ fontSize: '0.75rem', color: '#f59e0b', fontWeight: 700 }}>Planifié au calendrier</div>
        </div>
      </div>

      <div style={{ padding: '1.25rem 2.5rem', background: 'transparent', display: 'flex', gap: '15px' }}>
        <button className="btn-action" onClick={addAction} style={{ padding: '12px 28px', borderRadius: '30px', fontWeight: 800, display: 'flex', alignItems: 'center', gap: '10px', background: 'var(--main-surface)', border: '1px solid var(--main-border)', color: 'var(--main-text)', cursor: 'pointer' }}>
          <span style={{ fontSize: '1.2rem' }}>+</span> Nouvelle Action
        </button>
        <button className="btn-Excel" onClick={() => exportToCSV(actions, 'actions.csv')} style={{ background: 'var(--excel-color)', color: '#fff', border: 'none', padding: '12px 32px', borderRadius: '30px', fontWeight: 900, cursor: 'pointer', display: 'flex', alignItems: 'center', gap: '10px', boxShadow: '0 8px 16px rgba(70, 155, 111, 0.2)', textTransform: 'uppercase', fontSize: '0.75rem', letterSpacing: '0.05em' }}>
          📥 Export Excellence
        </button>
      </div>
      <div style={{ flex: 1, overflow: 'auto', padding: '2.5rem' }}>
        <div style={{ background: 'var(--main-surface)', borderRadius: '24px', border: '1px solid var(--main-border)', overflow: 'hidden', boxShadow: 'var(--card-shadow)' }}>
          <table className="editable-table" style={{ width: '100%', borderCollapse: 'separate', borderSpacing: 0 }}>
            <thead>
              <tr style={{ background: 'var(--excel-color)', color: '#fff' }}>
                <th style={{ padding: '20px 25px', width: '70px', textAlign: 'center', fontWeight: 900, fontSize: '0.7rem', textTransform: 'uppercase', letterSpacing: '0.1em' }}>ID</th>
                <th style={{ padding: '20px 15px', textAlign: 'left', fontWeight: 900, fontSize: '0.7rem', textTransform: 'uppercase', letterSpacing: '0.1em' }}>Action à mener / Origine</th>
                <th style={{ padding: '20px 15px', width: '250px', textAlign: 'left', fontWeight: 900, fontSize: '0.7rem', textTransform: 'uppercase', letterSpacing: '0.1em' }}>Propriétaire Responsable</th>
                <th style={{ padding: '20px 15px', width: '150px', textAlign: 'center', fontWeight: 900, fontSize: '0.7rem', textTransform: 'uppercase', letterSpacing: '0.1em' }}>Priorité</th>
                <th style={{ padding: '20px 25px', width: '200px', textAlign: 'center', fontWeight: 900, fontSize: '0.7rem', textTransform: 'uppercase', letterSpacing: '0.1em' }}>État d'avancement</th>
                <th style={{ padding: '20px 15px', width: '70px', textAlign: 'center', fontWeight: 900, fontSize: '0.7rem', textTransform: 'uppercase', letterSpacing: '0.1em' }}>Action</th>
              </tr>
            </thead>
            <tbody style={{ color: 'var(--main-text)' }}>
              {actions.map((a: any, idx: number) => (
                <tr key={a.id} style={{ borderBottom: '1px solid var(--main-border)', background: idx % 2 === 0 ? 'transparent' : 'rgba(0,0,0,0.01)', transition: 'all 0.2s ease' }}>
                  <td style={{ padding: '20px 25px', textAlign: 'center', color: 'var(--text-muted)', fontSize: '0.85rem', fontWeight: 900 }}>{a.id}</td>
                  <td style={{ padding: '20px 15px' }}>
                    <input value={a.action} onChange={e => handleChange(a.id, 'action', e.target.value)} style={{ width: '100%', border: 'none', background: 'transparent', outline: 'none', fontWeight: 800, color: 'var(--main-text)', fontSize: '0.95rem', marginBottom: '4px' }} />
                    <div style={{ fontSize: '0.75rem', color: 'var(--text-muted)', fontWeight: 500 }}>Origine: {a.origin}</div>
                  </td>
                  <td style={{ padding: '20px 15px' }}>
                    <div style={{ display: 'flex', alignItems: 'center', gap: '10px' }}>
                      <div style={{ width: '32px', height: '32px', borderRadius: '10px', background: 'var(--main-bg)', border: '1px solid var(--main-border)', display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: '0.8rem', fontWeight: 900, color: 'var(--primary)' }}>{a.resp.charAt(0)}</div>
                      <input value={a.resp} onChange={e => handleChange(a.id, 'resp', e.target.value)} style={{ flex: 1, border: 'none', background: 'transparent', outline: 'none', color: 'var(--main-text)', fontSize: '0.9rem', fontWeight: 700 }} />
                    </div>
                  </td>
                  <td style={{ padding: '20px 15px', textAlign: 'center' }}>
                    <span style={{ padding: '6px 14px', background: a.prio === 'P1' ? 'rgba(239, 68, 68, 0.1)' : 'rgba(59, 130, 246, 0.1)', color: a.prio === 'P1' ? '#ef4444' : '#3b82f6', borderRadius: '12px', fontSize: '0.75rem', fontWeight: 950, border: '1px solid currentColor' }}>{a.prio}</span>
                  </td>
                  <td style={{ padding: '20px 25px', textAlign: 'center' }}>
                    <div style={{
                      padding: '8px 18px',
                      borderRadius: '30px',
                      fontSize: '0.7rem',
                      fontWeight: 950,
                      background: a.etat === 'Réalisé' ? 'rgba(16, 185, 129, 0.12)' : a.etat === 'En cours' ? 'rgba(59, 130, 246, 0.12)' : 'rgba(148, 163, 184, 0.12)',
                      color: a.etat === 'Réalisé' ? '#10b981' : a.etat === 'En cours' ? '#3b82f6' : '#64748b',
                      border: '1px solid currentColor',
                      textTransform: 'uppercase',
                      letterSpacing: '0.05em',
                      display: 'inline-block'
                    }}>
                      <select value={a.etat} onChange={e => handleChange(a.id, 'etat', e.target.value)} style={{ border: 'none', background: 'transparent', color: 'inherit', fontWeight: 'inherit', outline: 'none', cursor: 'pointer' }}>
                        <option value="À faire">À faire</option>
                        <option value="En cours">En cours</option>
                        <option value="Réalisé">Réalisé</option>
                        <option value="Bloqué">Bloqué</option>
                      </select>
                    </div>
                  </td>
                  <td style={{ padding: '20px 15px', textAlign: 'center' }}>
                    <button
                      onClick={() => deleteAction(a.id)}
                      title="Supprimer cette action"
                      style={{ background: 'transparent', border: 'none', cursor: 'pointer', fontSize: '1.1rem', opacity: 0.7, transition: 'opacity 0.2s' }}
                      onMouseEnter={e => e.currentTarget.style.opacity = '1'}
                      onMouseLeave={e => e.currentTarget.style.opacity = '0.7'}
                    >
                      🗑️
                    </button>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
};

const MatriceExigencesView = ({ projectId, fileId, currentDossierId, page = 'option1' }: { projectId: string, fileId: string, currentDossierId: string | null, page?: string }) => {
  const metadata = { projectName: 'SAP30 / CMRL', manager: 'Mme Dupont', version: 'V2.3', lastUpdate: '15/05/2023' };
  const dossierSubKey = currentDossierId ? `-${currentDossierId}` : "";
  const storageKey = `matrice-exigences-${projectId}-${fileId}${dossierSubKey}`;

  if (page === 'option2') {
    return (
      <div style={{ flex: 1, display: 'flex', flexDirection: 'column', background: 'var(--main-bg)', color: 'var(--main-text)', padding: '2.5rem' }}>
        <ExecutionHeader title="06 — Couverture des Exigences" metadata={metadata} />
        <div style={{ background: 'var(--main-surface)', borderRadius: '24px', padding: '3rem', border: '1px solid var(--main-border)', textAlign: 'center' }}>
          <div style={{ fontSize: '3rem', marginBottom: '1.5rem' }}>📈</div>
          <h3 style={{ fontSize: '1.5rem', fontWeight: 900 }}>Taux de couverture des exigences</h3>
          <p style={{ color: 'var(--text-muted)' }}>Analyse graphique de la traçabilité entre exigences et livrables.</p>
        </div>
      </div>
    );
  }

  const defaultRows = [
    { id: 'E1', exig: 'Standard SAP', desc: "Le modèle standard SAP du siège sera la base du nouveau système", prio: 'Should have', type: 'Technique', just: 'Contrôle', livr: '-', dem: 'Directeur informatique', impact: 'Essentielle', acc: 'Oui', stat: 'En cours' },
    { id: 'E2', exig: 'Spécifiques CMRL', desc: "Des configurations spécifiques seront paramétrées si elles ont fait l'objet d'une approbation du comité de pilotage", prio: 'Could have', type: 'Technique', just: 'Satisfaction des besoins', livr: '-', dem: 'Sponsor', impact: 'Essentielle', acc: 'Non', stat: 'Terminée' },
    { id: 'E3', exig: 'Justification spécifiques CMRL', desc: "Pour information, ces spécifiques doivent justifier de raisons réglementaires ou légales...", prio: 'Could have', type: 'Technique', just: 'Satisfaction des besoins', livr: '-', dem: 'Sponsor', impact: 'Essentielle', acc: 'Oui', stat: 'Non commencée' },
    { id: 'E4', exig: 'Réseau informatique', desc: "Le réseau informatique sera adapté aux besoins en fonction de l'étude...", prio: 'Could have', type: 'Technique', just: 'Satisfaction des besoins', livr: 'L0 / Lot 1', dem: 'Chef de projet', impact: 'Utile', acc: 'Oui', stat: 'En cours' },
    { id: 'E5', exig: 'Coordination CMRL', desc: "Un responsable de l'intégration, monsieur Jacques Meunier, a été assigné...", prio: 'Must have', type: 'Fonctionnelle', just: 'Support', livr: '-', dem: 'Chef de projet', impact: 'Utile', acc: 'Oui', stat: 'Terminée' },
    { id: 'E6', exig: 'Conduite du changement', desc: "Un/une responsable du changement devra être assigné dans les jours qui suivent", prio: 'Must have', type: 'Fonctionnelle', just: 'Satisfaction des parties prenantes', livr: '-', dem: 'Chef de projet', impact: 'Utile', acc: 'Oui', stat: 'En cours' },
    { id: 'E7', exig: 'Ressources Informatiques', desc: "Les ressources informatiques seront assignées à temps plein au projet", prio: 'Should have', type: 'Fonctionnelle', just: 'Gestion des délais', livr: '-', dem: 'Chef de projet', impact: 'Essentielle', acc: 'Oui', stat: 'En cours' },
  ];

  const [rows, setRows] = useState(() => {
    const saved = localStorage.getItem(storageKey);
    return saved ? JSON.parse(saved) : defaultRows;
  });

  useEffect(() => {
    localStorage.setItem(storageKey, JSON.stringify(rows));
  }, [rows, storageKey]);

  const addExigence = () => {
    const nextNum = rows.length + 1;
    const newExig = {
      id: `E${nextNum}`,
      exig: `Nouvelle exigence ${nextNum}`,
      desc: 'Description détaillée du besoin fonctionnel ou technique...',
      prio: 'Should have',
      type: 'Fonctionnelle',
      just: 'Satisfaction des besoins',
      livr: '-',
      dem: 'Chef de projet',
      impact: 'Utile',
      acc: 'Oui',
      stat: 'Non commencée'
    };
    setRows([...rows, newExig]);
  };

  const deleteExigence = (id: string) => {
    setRows(rows.filter((r: any) => r.id !== id));
  };

  const handleChange = (id: string, field: string, val: any) => {
    setRows(rows.map((r: any) => r.id === id ? { ...r, [field]: val } : r));
  };

  const getPrioStyles = (prio: string) => {
    switch (prio) {
      case 'Must have': return { bg: 'rgba(239, 68, 68, 0.1)', color: '#ef4444', label: 'HAUTE' };
      case 'Should have': return { bg: 'rgba(59, 130, 246, 0.1)', color: '#3b82f6', label: 'MOYENNE' };
      case 'Could have': return { bg: 'rgba(16, 185, 129, 0.1)', color: '#10b981', label: 'BASSE' };
      default: return { bg: 'rgba(148, 163, 184, 0.1)', color: '#94a3b8', label: 'N/A' };
    }
  };

  const getStatusStyles = (stat: string) => {
    switch (stat) {
      case 'Terminée': return { bg: 'linear-gradient(135deg, #10b981, #059669)', color: '#fff' };
      case 'En cours': return { bg: 'linear-gradient(135deg, #3b82f6, #2563eb)', color: '#fff' };
      case 'Non commencée': return { bg: 'rgba(148, 163, 184, 0.2)', color: 'var(--text-muted)' };
      default: return { bg: 'transparent', color: 'inherit' };
    }
  };

  return (
    <div style={{ flex: 1, display: 'flex', flexDirection: 'column', background: 'var(--main-bg)', color: 'var(--main-text)', fontFamily: "'Outfit', sans-serif" }}>
      <ExecutionHeader title="06 — Matrice de Traçabilité des Exigences" metadata={metadata} />

      <div style={{
        padding: '1.5rem 3rem',
        background: 'rgba(255, 255, 255, 0.02)',
        backdropFilter: 'blur(10px)',
        borderBottom: '1px solid var(--main-border)',
        display: 'flex',
        gap: '20px',
        alignItems: 'center',
        justifyContent: 'space-between'
      }}>
        <div style={{ display: 'flex', alignItems: 'center', gap: '15px' }}>
          <div style={{
            width: '42px',
            height: '42px',
            borderRadius: '12px',
            background: 'linear-gradient(135deg, #3b82f6, #6366f1)',
            display: 'flex',
            alignItems: 'center',
            justifyContent: 'center',
            boxShadow: '0 8px 16px rgba(59, 130, 246, 0.2)'
          }}>
            <span style={{ fontSize: '1.2rem' }}>📋</span>
          </div>
          <div>
            <h3 style={{ margin: 0, fontSize: '1.1rem', fontWeight: 900, letterSpacing: '-0.01em' }}>REGISTRE DES EXIGENCES</h3>
            <p style={{ margin: 0, fontSize: '0.75rem', color: 'var(--text-muted)', fontWeight: 600 }}>Traçabilité & Priorisation du périmètre</p>
          </div>
        </div>

        <div style={{ display: 'flex', gap: '12px' }}>
          <button
            className="btn-action"
            onClick={addExigence}
            style={{
              padding: '10px 24px',
              borderRadius: '12px',
              fontWeight: 800,
              display: 'flex',
              alignItems: 'center',
              gap: '8px',
              background: 'var(--main-surface)',
              border: '1px solid var(--main-border)',
              color: 'var(--main-text)',
              cursor: 'pointer',
              fontSize: '0.8rem',
              transition: 'all 0.3s ease'
            }}
          >
            <span>➕</span> Ajouter une exigence
          </button>
          <button
            onClick={() => exportToCSV(rows, 'matrice_exigences.csv')}
            style={{
              padding: '10px 24px',
              background: 'var(--excel-color)',
              color: '#fff',
              border: 'none',
              borderRadius: '12px',
              cursor: 'pointer',
              fontWeight: 900,
              fontSize: '0.8rem',
              display: 'flex',
              alignItems: 'center',
              gap: '10px',
              boxShadow: '0 8px 16px rgba(16, 185, 129, 0.2)',
              transition: 'all 0.3s ease'
            }}
          >
            <span>📥</span> EXPORT SMART
          </button>
        </div>
      </div>

      <div style={{ flex: 1, padding: '2.5rem 3rem', overflow: 'auto' }}>
        <div style={{
          background: 'var(--main-surface)',
          borderRadius: '24px',
          border: '1px solid var(--main-border)',
          boxShadow: '0 20px 40px rgba(0,0,0,0.05)',
          overflow: 'hidden'
        }}>
          <table style={{ width: '100%', borderCollapse: 'separate', borderSpacing: 0 }}>
            <thead>
              <tr style={{ background: 'rgba(16, 185, 129, 0.03)' }}>
                <th style={{ padding: '24px 25px', width: '80px', textAlign: 'center', fontWeight: 900, fontSize: '0.7rem', textTransform: 'uppercase', color: 'var(--excel-color)', letterSpacing: '0.1em' }}>ID</th>
                <th style={{ padding: '24px 15px', textAlign: 'left', fontWeight: 900, fontSize: '0.7rem', textTransform: 'uppercase', color: 'var(--text-muted)', letterSpacing: '0.1em' }}>Exigence & Spécification</th>
                <th style={{ padding: '24px 15px', textAlign: 'center', fontWeight: 900, fontSize: '0.7rem', textTransform: 'uppercase', color: 'var(--text-muted)', letterSpacing: '0.1em', width: '150px' }}>Criticité</th>
                <th style={{ padding: '24px 15px', textAlign: 'left', fontWeight: 900, fontSize: '0.7rem', textTransform: 'uppercase', color: 'var(--text-muted)', letterSpacing: '0.1em', width: '200px' }}>Emetteur</th>
                <th style={{ padding: '24px 25px', textAlign: 'center', fontWeight: 900, fontSize: '0.7rem', textTransform: 'uppercase', color: 'var(--text-muted)', letterSpacing: '0.1em', width: '180px' }}>État</th>
                <th style={{ padding: '24px 15px', textAlign: 'center', fontWeight: 900, fontSize: '0.7rem', textTransform: 'uppercase', color: 'var(--text-muted)', letterSpacing: '0.1em', width: '70px' }}>Action</th>
              </tr>
            </thead>
            <tbody style={{ color: 'var(--main-text)' }}>
              {rows.map((r, i) => {
                const prio = getPrioStyles(r.prio);
                const stat = getStatusStyles(r.stat);
                return (
                  <tr key={i} className="premium-row" style={{
                    borderTop: '1px solid var(--main-border)',
                    background: 'transparent',
                    transition: 'all 0.2s ease'
                  }}>
                    <td style={{ padding: '25px', textAlign: 'center', fontWeight: 950, color: 'var(--excel-color)', fontSize: '0.95rem' }}>{r.id}</td>
                    <td style={{ padding: '25px 15px' }}>
                      <div style={{ fontWeight: 800, color: 'var(--main-text)', fontSize: '1rem', marginBottom: '6px' }}>{r.exig}</div>
                      <div style={{ fontSize: '0.85rem', color: 'var(--text-muted)', lineHeight: '1.5', fontWeight: 500, maxWidth: '600px' }}>{r.desc}</div>
                      <div style={{ marginTop: '12px', display: 'flex', gap: '10px' }}>
                        <span style={{ fontSize: '0.65rem', background: 'rgba(59, 130, 246, 0.05)', color: '#3b82f6', padding: '4px 10px', borderRadius: '6px', border: '1px solid rgba(59, 130, 246, 0.1)', fontWeight: 900, textTransform: 'uppercase' }}>TYPE: {r.type}</span>
                        <span style={{ fontSize: '0.65rem', background: 'rgba(16, 185, 129, 0.05)', color: '#10b981', padding: '4px 10px', borderRadius: '6px', border: '1px solid rgba(16, 185, 129, 0.1)', fontWeight: 900, textTransform: 'uppercase' }}>JUST: {r.just}</span>
                      </div>
                    </td>
                    <td style={{ padding: '25px 15px', textAlign: 'center' }}>
                      <div style={{
                        padding: '8px 16px',
                        borderRadius: '10px',
                        fontSize: '0.7rem',
                        fontWeight: 950,
                        background: prio.bg,
                        color: prio.color,
                        border: '1px solid currentColor',
                        display: 'inline-flex',
                        alignItems: 'center',
                        gap: '6px'
                      }}>
                        <div style={{ width: '6px', height: '6px', borderRadius: '50%', background: 'currentColor' }}></div>
                        {prio.label}
                      </div>
                    </td>
                    <td style={{ padding: '25px 15px' }}>
                      <div style={{ display: 'flex', alignItems: 'center', gap: '10px' }}>
                        <div style={{ width: '30px', height: '30px', borderRadius: '8px', background: 'var(--main-bg)', display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: '0.75rem', fontWeight: 900, border: '1px solid var(--main-border)', color: 'var(--primary)' }}>
                          {r.dem.charAt(0)}
                        </div>
                        <span style={{ fontWeight: 700, fontSize: '0.9rem' }}>{r.dem}</span>
                      </div>
                    </td>
                    <td style={{ padding: '25px', textAlign: 'center' }}>
                      <div style={{
                        padding: '10px 20px',
                        borderRadius: '30px',
                        fontSize: '0.7rem',
                        fontWeight: 950,
                        background: stat.bg,
                        color: stat.color,
                        boxShadow: r.stat !== 'Non commencée' ? '0 4px 12px rgba(0,0,0,0.1)' : 'none',
                        textTransform: 'uppercase',
                        letterSpacing: '0.05em',
                        display: 'inline-block'
                      }}>{r.stat}</div>
                    </td>
                    <td style={{ padding: '25px 15px', textAlign: 'center' }}>
                      <button
                        onClick={() => deleteExigence(r.id)}
                        title="Supprimer cette exigence"
                        style={{ background: 'transparent', border: 'none', cursor: 'pointer', fontSize: '1.1rem', opacity: 0.7, transition: 'opacity 0.2s' }}
                        onMouseEnter={e => e.currentTarget.style.opacity = '1'}
                        onMouseLeave={e => e.currentTarget.style.opacity = '0.7'}
                      >
                        🗑️
                      </button>
                    </td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      </div>
      <style>{`
        .premium-row:hover {
          background: rgba(59, 130, 246, 0.02) !important;
          transform: scale(1.002);
          box-shadow: inset 4px 0 0 var(--excel-color);
        }
      `}</style>
    </div>
  );
};

// ─── WBSTasksView (WBS SAP) ───────────────────────────────────────
const WBSTasksView = () => {
  const wbsData = [
    { lvl: 1, id: '1', act: 'Initialisation', charge: 48, duration: '20 j', pred: '-', prog: 51, bold: true },
    { lvl: 2, id: '1.1', act: "Etude d'opportunité et approbation", charge: 48, duration: '20 j', pred: '0', prog: 51 },
    { lvl: 2, id: '1.2', act: 'Préparation de la Charte Projet - Justification', charge: 4, duration: '4 j', pred: '1.0', prog: 30 },
    { lvl: 2, id: '1.3', act: 'Estimation coûts/délais fonctionnels Finances', charge: 5, duration: '5 j', pred: '1.1', prog: 50 },
    { lvl: 2, id: '1.4', act: 'Estimation coûts/délais fonctionnels Ventes', charge: 1, duration: '1 j', pred: '1.1', prog: 60 },
    { lvl: 2, id: '1.5', act: 'Estimation coûts/délais fonctionnels Achats', charge: 2, duration: '2 j', pred: '1.1', prog: 10 },
    { lvl: 2, id: '1.6', act: 'Estimation coûts/délais fonctionnels Logistique', charge: 5, duration: '5 j', pred: '1.1', prog: 20 },
    { lvl: 2, id: '1.7', act: 'Estimation délais/ressources Infrastructure', charge: 5, duration: '5 j', pred: '1.1', prog: 20 },
    { lvl: 2, id: '1.8', act: 'Estimation coûts/délais fonctionnels Production', charge: 5, duration: '5 j', pred: '1.1', prog: 20 },
    { lvl: 2, id: '1.9', act: 'Rédaction de la Charte Projet', charge: 15, duration: '15 j', pred: '1.0', prog: 15 },
    { lvl: 2, id: '1.10', act: 'Approbation de la Charte Projet', charge: 20, duration: '20 j', pred: '1.0', prog: 100 },
    { lvl: 1, id: '2', act: 'Planification', charge: 237, duration: '84 j', pred: '1.0', prog: 58, bold: true },
    { lvl: 2, id: '2.1', act: 'Conception générale Financière', charge: 35, duration: '35 j', pred: '-', prog: 50 },
    { lvl: 2, id: '2.2', act: 'Rédaction du cahier des charges', charge: 50, duration: '50 j', pred: '1.5', prog: 90 },
    { lvl: 1, id: '3', act: 'Réalisation', charge: 504, duration: '180 j', pred: '2.2-2.4', prog: 31, bold: true },
    { lvl: 2, id: '3.1', act: 'Paramétrage et tests unitaires', charge: 65, duration: '20 j', pred: '', prog: 20 },
  ];
  return (
    <div style={{ flex: 1, overflow: 'auto', background: '#fff', color: '#000', padding: '1.5rem' }}>
      <div style={{ display: 'flex', alignItems: 'center', marginBottom: '1rem', gap: '0.8rem', justifyContent: 'space-between' }}>
        <div style={{ display: 'flex', alignItems: 'center', gap: '0.8rem' }}>
          <div style={{ width: '28px', height: '28px', borderRadius: '50%', background: '#4ada8e', display: 'flex', alignItems: 'center', justifyContent: 'center' }}><span style={{ color: '#fff', fontWeight: 'bold' }}>⟡</span></div>
          <div style={{ background: '#f97316', color: '#fff', padding: '8px 16px', borderRadius: '4px', fontWeight: 'bold' }}>WBS SAP</div>
        </div>
        <button
          onClick={() => exportToCSV(wbsData, 'wbs_sap.csv')}
          style={{ padding: '8px 16px', background: '#3b82f6', color: '#fff', border: 'none', borderRadius: '4px', cursor: 'pointer', fontWeight: 'bold', fontSize: '0.8rem' }}
        >
          📥 Export CSV
        </button>
      </div>
      <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '0.75rem' }}>
        <thead><tr style={{ background: '#469b6f', color: '#fff' }}>
          <th style={{ padding: '7px', border: '1px solid #fff', width: '45px' }}>Niveau</th>
          <th style={{ padding: '7px', border: '1px solid #fff', width: '45px' }}>ID</th>
          <th style={{ padding: '7px', border: '1px solid #fff', textAlign: 'left' }}>WBS SAP</th>
          <th style={{ padding: '7px', border: '1px solid #fff', width: '80px' }}>Charge (h)</th>
          <th style={{ padding: '7px', border: '1px solid #fff', width: '80px' }}>Durée</th>
          <th style={{ padding: '7px', border: '1px solid #fff', width: '80px' }}>Prédécesseur</th>
          <th style={{ padding: '7px', border: '1px solid #fff', width: '80px' }}>Avancement (%)</th>
        </tr></thead>
        <tbody>{wbsData.map((r, i) => (
          <tr key={i} style={{ background: r.bold ? '#cbe4d5' : i % 2 === 0 ? '#fff' : '#f8fafc', borderBottom: '1px solid #eee', fontWeight: r.bold ? 'bold' : 'normal' }}>
            <td style={{ padding: '6px', border: '1px solid #ddd', textAlign: 'center', color: r.bold ? '#166534' : '#000' }}>{r.lvl}</td>
            <td style={{ padding: '6px', border: '1px solid #ddd', textAlign: 'center', color: r.bold ? '#166534' : '#469b6f' }}>{r.id}</td>
            <td style={{ padding: '6px', border: '1px solid #ddd', paddingLeft: r.bold ? '8px' : '20px', color: r.bold ? '#166534' : '#000' }}>{r.act}</td>
            <td style={{ padding: '6px', border: '1px solid #ddd', textAlign: 'center' }}>{r.charge}</td>
            <td style={{ padding: '6px', border: '1px solid #ddd', textAlign: 'center' }}>{r.duration}</td>
            <td style={{ padding: '6px', border: '1px solid #ddd', textAlign: 'center' }}>{r.pred}</td>
            <td style={{ padding: '6px', border: '1px solid #ddd', textAlign: 'center' }}>{r.prog}%</td>
          </tr>
        ))}</tbody>
      </table>
    </div>
  );
};

// ─── SuiviBudgetView ─────────────────────────────────────────────
const SuiviBudgetView = ({ page, projectId, fileId, currentDossierId }: { page: 'vp' | 'cr' | 'va' | 'rapport' | 'avanc', projectId: string, fileId: string, currentDossierId: string | null }) => {
  const metadata = { projectName: 'SAP30 / CMRL', manager: 'Mme Dupont', version: 'V2.3', lastUpdate: '28/11/2022' };
  const dossierSubKey = currentDossierId ? `-${currentDossierId}` : "";
  const storageKey = `faab-budget-data-${projectId}-${fileId}${dossierSubKey}`;

  const [budgetData, setBudgetData] = useState(() => {
    const saved = localStorage.getItem(storageKey);
    if (saved) return JSON.parse(saved);
    return [
      { sdp: 1, task: 'Expression des besoins', total: 19500, vp: 4000, cr: 4200, va: 4000 },
      { sdp: 2, task: 'Planification', total: 36000, vp: 12000, cr: 11500, va: 12000 },
      { sdp: 12, task: 'Clôture novembre 2022', total: 17000, vp: 12000, cr: 12800, va: 12000 },
    ];
  });

  useEffect(() => {
    localStorage.setItem(storageKey, JSON.stringify(budgetData));
  }, [budgetData]);

  const handleChange = (sdp: number, field: string, val: any) => {
    setBudgetData(budgetData.map((b: any) => b.sdp === sdp ? { ...b, [field]: parseFloat(val) || 0 } : b));
  };

  if (page === 'avanc') {
    return (
      <div style={{ flex: 1, display: 'flex', flexDirection: 'column', background: 'var(--main-bg)' }}>
        <ExecutionHeader title="28 — Courbe d'avancement" metadata={metadata} />
        <div style={{ padding: '2.5rem', flex: 1 }}>
          <div style={{ background: 'var(--main-surface)', borderRadius: '16px', border: '1px solid var(--main-border)', padding: '2.5rem', height: '100%', display: 'flex', flexDirection: 'column', boxShadow: 'var(--card-shadow)' }}>
            <h3 style={{ color: 'var(--main-text)', marginBottom: '2.5rem', fontWeight: 800, fontSize: '1.25rem' }}>📈 Performance Budgétaire Immédiate</h3>
            <div style={{ flex: 1, position: 'relative', borderLeft: '2px solid var(--main-border)', borderBottom: '2px solid var(--main-border)', margin: '2.5rem', background: 'rgba(59, 130, 246, 0.02)' }}>
              <svg width="100%" height="100%" viewBox="0 0 1000 400" preserveAspectRatio="none">
                <defs>
                  <linearGradient id="lineGrad" x1="0" y1="0" x2="0" y2="1">
                    <stop offset="0%" stopColor="rgba(59, 130, 246, 0.2)" />
                    <stop offset="100%" stopColor="rgba(59, 130, 246, 0)" />
                  </linearGradient>
                </defs>
                <path d="M0,350 L100,320 L200,300 L300,250 L400,200 L500,150 L600,100 L700,80 L800,50 L900,20" fill="url(#lineGrad)" />
                <path d="M0,350 L100,320 L200,300 L300,250 L400,200 L500,150 L600,100 L700,80 L800,50 L900,20" fill="none" stroke="#3b82f6" strokeWidth="5" strokeLinecap="round" />
                <path d="M0,350 L100,340 L200,320 L300,280 L400,250 L500,210 L600,180 L700,150" fill="none" stroke="#f59e0b" strokeWidth="4" strokeDasharray="10,6" />
              </svg>
              <div style={{ position: 'absolute', bottom: '-25px', width: '100%', display: 'flex', justifyContent: 'space-between', fontSize: '0.75rem', color: 'var(--text-muted)', fontWeight: 700 }}>
                <span>MARS</span><span>MAI</span><span>JUILLET</span><span>SEPT</span><span>NOV</span><span>JANV</span>
              </div>
            </div>
            <div style={{ display: 'flex', gap: '3rem', justifyContent: 'center', marginTop: '1.5rem', background: 'var(--main-bg)', padding: '15px', borderRadius: '12px', alignSelf: 'center', border: '1px solid var(--main-border)' }}>
              <div style={{ display: 'flex', alignItems: 'center', gap: '10px' }}>
                <div style={{ width: '16px', height: '16px', borderRadius: '4px', background: '#3b82f6' }}></div>
                <span style={{ fontSize: '0.85rem', fontWeight: 700, color: 'var(--main-text)' }}>Planifié (VP)</span>
              </div>
              <div style={{ display: 'flex', alignItems: 'center', gap: '10px' }}>
                <div style={{ width: '16px', height: '4px', background: '#f59e0b' }}></div>
                <span style={{ fontSize: '0.85rem', fontWeight: 700, color: 'var(--main-text)' }}>Réel (CR)</span>
              </div>
            </div>
          </div>
        </div>
      </div>
    );
  }

  if (page === 'rapport') {
    const totalCR = budgetData.reduce((acc: number, b: any) => acc + b.cr, 0);
    const totalVA = budgetData.reduce((acc: number, b: any) => acc + b.va, 0);
    const totalVP = budgetData.reduce((acc: number, b: any) => acc + b.vp, 0);
    const cpi = totalCR > 0 ? totalVA / totalCR : 1;
    const spi = totalVP > 0 ? totalVA / totalVP : 1;

    return (
      <div style={{ flex: 1, display: 'flex', flexDirection: 'column', background: 'var(--main-bg)' }}>
        <ExecutionHeader title="28 — Rapport VA (Earned Value)" metadata={metadata} />
        <div style={{ padding: '2.5rem' }}>
          <div style={{ display: 'grid', gridTemplateColumns: 'repeat(3, 1fr)', gap: '2rem' }}>
            <div style={{ background: 'var(--main-surface)', padding: '2.5rem', borderRadius: '16px', border: '1px solid var(--main-border)', boxShadow: 'var(--card-shadow)', textAlign: 'center' }}>
              <div style={{ color: 'var(--text-muted)', fontSize: '0.8rem', fontWeight: 800, textTransform: 'uppercase', marginBottom: '1rem' }}>Indice de Performance Coût (CPI)</div>
              <div style={{ fontSize: '3.5rem', fontWeight: 950, color: cpi >= 1 ? '#10b981' : '#ef4444', marginBottom: '12px' }}>{cpi.toFixed(2)}</div>
              <div style={{ display: 'inline-block', padding: '8px 16px', borderRadius: '30px', background: cpi >= 1 ? 'rgba(16, 185, 129, 0.12)' : 'rgba(239, 68, 68, 0.12)', color: cpi >= 1 ? '#10b981' : '#ef4444', fontSize: '0.75rem', fontWeight: 900 }}>
                {cpi >= 1 ? 'MAÎTRISE DES COÛTS' : 'SURCOÛT DÉTECTÉ'}
              </div>
            </div>
            <div style={{ background: 'var(--main-surface)', padding: '2.5rem', borderRadius: '16px', border: '1px solid var(--main-border)', boxShadow: 'var(--card-shadow)', textAlign: 'center' }}>
              <div style={{ color: 'var(--text-muted)', fontSize: '0.8rem', fontWeight: 800, textTransform: 'uppercase', marginBottom: '1rem' }}>Indice de Performance Délais (SPI)</div>
              <div style={{ fontSize: '3.5rem', fontWeight: 950, color: spi >= 1 ? '#10b981' : '#f59e0b', marginBottom: '12px' }}>{spi.toFixed(2)}</div>
              <div style={{ display: 'inline-block', padding: '8px 16px', borderRadius: '30px', background: spi >= 1 ? 'rgba(16, 185, 129, 0.12)' : 'rgba(245, 158, 11, 0.12)', color: spi >= 1 ? '#10b981' : '#f59e0b', fontSize: '0.75rem', fontWeight: 900 }}>
                {spi >= 1 ? 'DANS LES DÉLAIS' : 'RETARD PLANNING'}
              </div>
            </div>
            <div style={{ background: 'var(--main-surface)', padding: '2.5rem', borderRadius: '16px', border: '1px solid var(--main-border)', boxShadow: 'var(--card-shadow)', textAlign: 'center', display: 'flex', flexDirection: 'column', justifyContent: 'center' }}>
              <div style={{ color: 'var(--text-muted)', fontSize: '0.8rem', fontWeight: 800, textTransform: 'uppercase', marginBottom: '0.5rem' }}>Prévisionnel Terminaison (EAC)</div>
              <div style={{ fontSize: '2.8rem', fontWeight: 950, color: 'var(--main-text)' }}>{(1150000 / (cpi || 1)).toLocaleString()} $</div>
              <div style={{ fontSize: '0.85rem', color: 'var(--text-muted)', fontStyle: 'italic', marginTop: '10px' }}>Basé sur l'efficience actuelle</div>
            </div>
          </div>
        </div>
      </div>
    );
  }

  const months = ['mars-22', 'avr-22', 'mai-22', 'juin-22', 'juil-22', 'août-22', 'sept-22', 'oct-22', 'nov-22', 'déc-22', 'janv-23', 'févr-23'];

  const getTheme = () => {
    if (page === 'cr') return { color: '#f59e0b', label: 'Consommation Réelle (CR)', icon: '📉', field: 'cr' };
    if (page === 'vp') return { color: '#3b82f6', label: 'Valeur Planifiée (VP)', icon: '💰', field: 'vp' };
    if (page === 'va') return { color: '#10b981', label: 'Valeur Acquise (VA)', icon: '🟢', field: 'va' };
    return { color: '#64748b', label: 'Budget', icon: '📊', field: 'vp' };
  };

  const theme = getTheme();

  return (
    <div style={{ flex: 1, display: 'flex', flexDirection: 'column', background: 'var(--main-bg)' }}>
      <ExecutionHeader title={`28 — Suivi de budget - ${theme.label}`} metadata={metadata} />
      <div style={{ padding: '1.25rem 2.5rem', background: 'var(--main-surface)', borderBottom: '1px solid var(--main-border)', display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
        <div style={{ display: 'inline-flex', alignItems: 'center', gap: '10px', background: `${theme.color}15`, color: theme.color, padding: '8px 18px', borderRadius: '30px', fontSize: '0.75rem', fontWeight: 900, textTransform: 'uppercase', letterSpacing: '0.05em' }}>
          <span style={{ fontSize: '1.1rem' }}>{theme.icon}</span> VUE : {theme.label}
        </div>
        <button onClick={() => exportToCSV(budgetData, `budget_${page}.csv`)} className="btn-Excel" style={{ background: theme.color, color: '#fff', border: 'none', padding: '10px 24px', borderRadius: '10px', fontWeight: 800, cursor: 'pointer', display: 'flex', alignItems: 'center', gap: '8px' }}>
          <span>📥</span> Export CSV
        </button>
      </div>
      <div style={{ flex: 1, overflow: 'auto', padding: '2.5rem' }}>
        <div style={{ background: 'var(--main-surface)', borderRadius: '16px', border: '1px solid var(--main-border)', overflow: 'hidden', boxShadow: 'var(--card-shadow)' }}>
          <table className="editable-table" style={{ width: '100%', borderCollapse: 'separate', borderSpacing: 0 }}>
            <thead>
              <tr style={{ background: theme.color, color: '#fff', fontWeight: 800 }}>
                <th style={{ padding: '18px 15px', width: '80px', textAlign: 'center' }}>WBS</th>
                <th style={{ padding: '18px 15px', textAlign: 'left' }}>Désignation de la Tâche</th>
                <th style={{ padding: '18px 15px', width: '180px', textAlign: 'right' }}>Total Budget ($)</th>
                <th style={{ padding: '18px 15px', width: '180px', textAlign: 'right' }}>{theme.label} ($)</th>
                <th style={{ padding: '18px 15px', width: '150px', textAlign: 'right' }}>Écart Final</th>
              </tr>
            </thead>
            <tbody style={{ color: 'var(--main-text)' }}>
              {budgetData.map((b: any, idx: number) => (
                <tr key={b.sdp} style={{ borderBottom: '1px solid var(--main-border)', background: idx % 2 === 0 ? 'transparent' : 'rgba(0,0,0,0.02)' }}>
                  <td style={{ padding: '15px', textAlign: 'center', color: 'var(--text-muted)', fontWeight: 800 }}>{b.sdp}</td>
                  <td style={{ padding: '15px' }}>
                    <div style={{ fontWeight: 700, fontSize: '0.95rem' }}>{b.task}</div>
                  </td>
                  <td style={{ padding: '15px', textAlign: 'right' }}>
                    <input type="number" value={b.total} onChange={e => handleChange(b.sdp, 'total', e.target.value)} style={{ width: '120px', border: '1px solid var(--main-border)', borderRadius: '6px', textAlign: 'right', fontWeight: 800, padding: '6px', background: 'var(--main-bg)', color: 'var(--main-text)' }} />
                  </td>
                  <td style={{ padding: '15px', textAlign: 'right' }}>
                    <input type="number" value={b[theme.field]} onChange={e => handleChange(b.sdp, theme.field, e.target.value)} style={{ width: '120px', border: '1px solid var(--main-border)', borderRadius: '6px', textAlign: 'right', fontWeight: 800, padding: '6px', background: 'var(--main-bg)', color: theme.color }} />
                  </td>
                  <td style={{ padding: '15px', textAlign: 'right' }}>
                    <div style={{ color: (b.total - b.cr) < 0 ? '#ef4444' : '#10b981', fontWeight: 950, fontSize: '1rem' }}>
                      {(b.total - b.cr).toLocaleString()} $
                    </div>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
};

// ─── ControleDashboardView ───────────────────────────────────────
const ControleDashboardView = ({ page = 'option1', projectId, fileId, currentDossierId }: { page?: string, projectId: string, fileId: string, currentDossierId: string | null }) => {
  const metadata = { projectName: 'FAAB CORE', manager: 'Oliver Martin', version: 'V3.4', lastUpdate: '28/11/2022' };
  const dossierSubKey = currentDossierId ? `-${currentDossierId}` : "";
  const storageKey = `faab-ctrl-dash-${projectId}-${fileId}${dossierSubKey}`;

  if (page === 'option2') {
    return (
      <div style={{ flex: 1, display: 'flex', flexDirection: 'column', background: 'var(--main-bg)', color: 'var(--main-text)' }}>
        <ExecutionHeader title="29 — Cosy Tracking Dashboard" metadata={metadata} />
        <div style={{ padding: '1.25rem 2.5rem', background: 'var(--main-surface)', borderBottom: '1px solid var(--main-border)', display: 'flex', gap: '10px' }}>
          <div style={{ display: 'inline-flex', alignItems: 'center', gap: '10px', background: 'rgba(59, 130, 246, 0.12)', color: '#3b82f6', padding: '8px 18px', borderRadius: '30px', fontSize: '0.75rem', fontWeight: 900, textTransform: 'uppercase', letterSpacing: '0.05em' }}>
            📈 VUE ACTIVE : COSY TRACKING
          </div>
        </div>
        <div style={{ flex: 1, padding: '2.5rem', display: 'flex', flexDirection: 'column', alignItems: 'center', justifyContent: 'center' }}>
          <div style={{ textAlign: 'center', maxWidth: '1000px', width: '100%' }}>
            <h2 style={{ fontSize: '2rem', fontWeight: 900, background: 'linear-gradient(135deg, #3b82f6, #2563eb)', WebkitBackgroundClip: 'text', WebkitTextFillColor: 'transparent', marginBottom: '10px' }}>📈 Suivi Cosy TracKing</h2>
            <p style={{ color: 'var(--text-muted)', fontSize: '1.1rem', marginBottom: '3rem', fontWeight: 500 }}>Performance opérationnelle et maintenance en temps réel</p>

            <div style={{ display: 'grid', gridTemplateColumns: 'repeat(3, 1fr)', gap: '2rem' }}>
              <div style={{ background: 'var(--main-surface)', padding: '2.5rem', borderRadius: '20px', border: '1px solid var(--main-border)', boxShadow: 'var(--card-shadow)', position: 'relative', overflow: 'hidden' }}>
                <div style={{ position: 'absolute', top: 0, left: 0, width: '4px', height: '100%', background: '#3b82f6' }}></div>
                <div style={{ color: 'var(--text-muted)', fontSize: '0.8rem', fontWeight: 800, textTransform: 'uppercase', marginBottom: '1rem' }}>Tickets Ouverts</div>
                <div style={{ fontSize: '3.5rem', fontWeight: 950, color: 'var(--main-text)' }}>12</div>
                <div style={{ color: '#3b82f6', fontSize: '0.85rem', fontWeight: 700, marginTop: '10px' }}>-2 depuis hier</div>
              </div>

              <div style={{ background: 'var(--main-surface)', padding: '2.5rem', borderRadius: '20px', border: '1px solid var(--main-border)', boxShadow: 'var(--card-shadow)', position: 'relative', overflow: 'hidden' }}>
                <div style={{ position: 'absolute', top: 0, left: 0, width: '4px', height: '100%', background: '#f59e0b' }}></div>
                <div style={{ color: 'var(--text-muted)', fontSize: '0.8rem', fontWeight: 800, textTransform: 'uppercase', marginBottom: '1rem' }}>Temps de Résolution</div>
                <div style={{ fontSize: '3.5rem', fontWeight: 950, color: 'var(--main-text)' }}>4.2h</div>
                <div style={{ color: '#f59e0b', fontSize: '0.85rem', fontWeight: 700, marginTop: '10px' }}>Cible: 6h max</div>
              </div>

              <div style={{ background: 'var(--main-surface)', padding: '2.5rem', borderRadius: '20px', border: '1px solid var(--main-border)', boxShadow: 'var(--card-shadow)', position: 'relative', overflow: 'hidden' }}>
                <div style={{ position: 'absolute', top: 0, left: 0, width: '4px', height: '100%', background: '#10b981' }}></div>
                <div style={{ color: 'var(--text-muted)', fontSize: '0.8rem', fontWeight: 800, textTransform: 'uppercase', marginBottom: '1rem' }}>SLA Respecté</div>
                <div style={{ fontSize: '3.5rem', fontWeight: 950, color: '#10b981' }}>98%</div>
                <div style={{ color: '#10b981', fontSize: '0.85rem', fontWeight: 700, marginTop: '10px' }}>Objectif critique : 95%</div>
              </div>
            </div>
          </div>
        </div>
      </div>
    );
  }

  if (page === 'option3') {
    return (
      <div style={{ flex: 1, display: 'flex', flexDirection: 'column', background: 'var(--main-bg)' }}>
        <ExecutionHeader title="29 — Chart GANTT (Revue)" metadata={metadata} />
        <div style={{ padding: '1.25rem 2.5rem', background: 'var(--main-surface)', borderBottom: '1px solid var(--main-border)', display: 'flex', gap: '10px' }}>
          <div style={{ display: 'inline-flex', alignItems: 'center', gap: '8px', background: 'rgba(16, 185, 129, 0.12)', color: '#10b981', padding: '8px 18px', borderRadius: '30px', fontSize: '0.75rem', fontWeight: 900, textTransform: 'uppercase', letterSpacing: '0.05em' }}>
            🕒 VUE ACTIVE : GANTT CHART
          </div>
        </div>
        <div style={{ padding: '2.5rem', flex: 1 }}>
          <div style={{ background: 'var(--main-surface)', borderRadius: '20px', border: '1px solid var(--main-border)', padding: '2.5rem', height: '100%', boxShadow: 'var(--card-shadow)' }}>
            <h3 style={{ marginBottom: '2.5rem', color: 'var(--main-text)', fontWeight: 800, fontSize: '1.1rem', textTransform: 'uppercase' }}>🕒 Timeline du Projet (Consolidée)</h3>
            <div style={{ display: 'flex', flexDirection: 'column', gap: '2rem' }}>
              {[
                { task: 'Initialisation', progress: 100, color: '#10b981', date: 'Terminé' },
                { task: 'Conception', progress: 85, color: '#3b82f6', date: 'En cours' },
                { task: 'Développement', progress: 40, color: '#f59e0b', date: 'Oct 2022' },
                { task: 'Recette', progress: 0, color: '#e2e8f0', date: 'Nov 2022' }
              ].map(t => (
                <div key={t.task}>
                  <div style={{ display: 'flex', justifyContent: 'space-between', marginBottom: '12px', fontSize: '0.9rem' }}>
                    <div style={{ display: 'flex', alignItems: 'center', gap: '10px' }}>
                      <span style={{ fontWeight: 800, color: 'var(--main-text)' }}>{t.task}</span>
                      <span style={{ fontSize: '0.75rem', color: 'var(--text-muted)', fontWeight: 600, padding: '2px 8px', background: 'var(--main-bg)', borderRadius: '4px' }}>{t.date}</span>
                    </div>
                    <span style={{ fontWeight: 900, color: t.color === '#e2e8f0' ? 'var(--text-muted)' : t.color }}>{t.progress}%</span>
                  </div>
                  <div style={{ height: '14px', background: 'var(--main-bg)', borderRadius: '7px', border: '1px solid var(--main-border)', overflow: 'hidden', padding: '2px' }}>
                    <div style={{ height: '100%', width: `${t.progress}%`, background: t.color, borderRadius: '5px' }}></div>
                  </div>
                </div>
              ))}
            </div>
          </div>
        </div>
      </div>
    );
  }

  if (page === 'option4') {
    return (
      <div style={{ flex: 1, display: 'flex', flexDirection: 'column', background: 'var(--main-bg)' }}>
        <ExecutionHeader title="29 — Time Budget / Ressources" metadata={metadata} />
        <div style={{ padding: '1.25rem 2.5rem', background: 'var(--main-surface)', borderBottom: '1px solid var(--main-border)', display: 'flex', gap: '10px' }}>
          <div style={{ display: 'inline-flex', alignItems: 'center', gap: '8px', background: 'rgba(245, 158, 11, 0.12)', color: '#f59e0b', padding: '8px 18px', borderRadius: '30px', fontSize: '0.75rem', fontWeight: 900, textTransform: 'uppercase', letterSpacing: '0.05em' }}>
            💰 VUE ACTIVE : BUDGET TEMPS
          </div>
        </div>
        <div style={{ padding: '2.5rem' }}>
          <div style={{ background: 'var(--main-surface)', borderRadius: '16px', border: '1px solid var(--main-border)', overflow: 'hidden', boxShadow: 'var(--card-shadow)' }}>
            <table className="editable-table" style={{ width: '100%', borderCollapse: 'separate', borderSpacing: 0 }}>
              <thead style={{ background: 'var(--excel-color)', color: '#fff' }}>
                <tr style={{ fontWeight: 800, textTransform: 'uppercase', fontSize: '0.75rem' }}>
                  <th style={{ padding: '18px 25px', textAlign: 'left' }}>Collaborateur / Ressource</th>
                  <th style={{ padding: '18px 25px', textAlign: 'center' }}>Heures Prévues</th>
                  <th style={{ padding: '18px 25px', textAlign: 'center' }}>Heures Réelles</th>
                  <th style={{ padding: '18px 25px', textAlign: 'center' }}>Écart (Delta)</th>
                </tr>
              </thead>
              <tbody style={{ color: 'var(--main-text)' }}>
                <tr style={{ borderBottom: '1px solid var(--main-border)' }}>
                  <td style={{ padding: '18px 25px', fontWeight: 700 }}>Oliver Martin</td>
                  <td style={{ padding: '18px 25px', textAlign: 'center', fontWeight: 800 }}>120h</td>
                  <td style={{ padding: '18px 25px', textAlign: 'center', fontWeight: 800, color: '#ef4444' }}>145h</td>
                  <td style={{ padding: '18px 25px', textAlign: 'center' }}>
                    <span style={{ padding: '6px 12px', background: 'rgba(239, 68, 68, 0.12)', color: '#ef4444', borderRadius: '30px', fontSize: '0.75rem', fontWeight: 900 }}>+25h EXTRAS</span>
                  </td>
                </tr>
                <tr style={{ borderBottom: '1px solid var(--main-border)', background: 'rgba(0,0,0,0.01)' }}>
                  <td style={{ padding: '18px 25px', fontWeight: 700 }}>Chantal Dupont</td>
                  <td style={{ padding: '18px 25px', textAlign: 'center', fontWeight: 800 }}>80h</td>
                  <td style={{ padding: '18px 25px', textAlign: 'center', fontWeight: 800, color: '#10b981' }}>72h</td>
                  <td style={{ padding: '18px 25px', textAlign: 'center' }}>
                    <span style={{ padding: '6px 12px', background: 'rgba(16, 185, 129, 0.12)', color: '#10b981', borderRadius: '30px', fontSize: '0.75rem', fontWeight: 900 }}>-8h OPTIMISÉ</span>
                  </td>
                </tr>
              </tbody>
            </table>
          </div>
        </div>
      </div>
    );
  }

  return (
    <div style={{ flex: 1, display: 'flex', flexDirection: 'column', background: 'var(--main-bg)', color: 'var(--main-text)' }}>
      <ExecutionHeader title="29 — Tableau de bord global" metadata={metadata} />
      <div style={{ padding: '1.25rem 2.5rem', background: 'var(--main-surface)', borderBottom: '1px solid var(--main-border)', display: 'flex', gap: '10px' }}>
        <div style={{ display: 'inline-flex', alignItems: 'center', gap: '10px', background: 'rgba(16, 185, 129, 0.12)', color: '#10b981', padding: '8px 18px', borderRadius: '30px', fontSize: '0.75rem', fontWeight: 900, textTransform: 'uppercase', letterSpacing: '0.05em' }}>
          📊 VUE ACTIVE : DASHBOARD GLOBAL
        </div>
      </div>

      <div style={{ padding: '2.5rem', display: 'grid', gridTemplateColumns: 'repeat(4, 1fr)', gap: '2rem' }}>
        <div style={{ background: 'var(--main-surface)', padding: '2rem', borderRadius: '16px', border: '1px solid var(--main-border)', borderLeft: '5px solid #10b981', boxShadow: 'var(--card-shadow)' }}>
          <div style={{ color: 'var(--text-muted)', fontSize: '0.75rem', fontWeight: 800, textTransform: 'uppercase' }}>Estimation Fin Projet</div>
          <div style={{ fontSize: '1.4rem', fontWeight: 950, marginTop: '8px' }}>Ré-estimé à ce jour</div>
          <div style={{ color: '#10b981', fontSize: '0.85rem', fontWeight: 700, marginTop: '4px' }}>Maintien du focus initial</div>
        </div>

        <div style={{ background: 'var(--main-surface)', padding: '2rem', borderRadius: '16px', border: '1px solid var(--main-border)', boxShadow: 'var(--card-shadow)', textAlign: 'center', display: 'flex', flexDirection: 'column', justifyContent: 'center' }}>
          <div style={{ fontSize: '2.8rem', fontWeight: 950, color: '#3b82f6' }}>68%</div>
          <div style={{ color: 'var(--text-muted)', fontSize: '0.75rem', fontWeight: 800, textTransform: 'uppercase' }}>Taux d'Avancement Global</div>
        </div>

        <div style={{ background: 'var(--main-surface)', padding: '2rem', borderRadius: '16px', border: '1px solid var(--main-border)', borderLeft: '5px solid #3b82f6', boxShadow: 'var(--card-shadow)' }}>
          <div style={{ color: 'var(--text-muted)', fontSize: '0.75rem', fontWeight: 800, textTransform: 'uppercase' }}>Budget Total Initial</div>
          <div style={{ fontSize: '1.8rem', fontWeight: 950, marginTop: '8px' }}>1,250,000 $</div>
        </div>

        <div style={{ background: 'linear-gradient(135deg, #f97316, #ea580c)', padding: '2rem', borderRadius: '16px', color: '#fff', boxShadow: '0 10px 15px -3px rgba(249, 115, 22, 0.3)' }}>
          <div style={{ opacity: 0.9, fontSize: '0.75rem', fontWeight: 800, textTransform: 'uppercase' }}>Budget Consommé</div>
          <div style={{ fontSize: '1.8rem', fontWeight: 950, marginTop: '8px' }}>845,000 $</div>
        </div>
      </div>

      <div style={{ display: 'grid', gridTemplateColumns: '1.1fr 0.9fr', gap: '2rem', padding: '0 2.5rem 2.5rem', flex: 1, overflow: 'auto' }}>
        <div style={{ background: 'var(--main-surface)', padding: '2.5rem', borderRadius: '20px', border: '1px solid var(--main-border)', boxShadow: 'var(--card-shadow)' }}>
          <h3 style={{ color: '#10b981', marginBottom: '2rem', fontSize: '1rem', fontWeight: 900, textTransform: 'uppercase' }}>🩺 Santé du Projet</h3>
          <table style={{ width: '100%', fontSize: '0.9rem', borderCollapse: 'collapse' }}>
            <tbody>
              <tr style={{ borderBottom: '1px solid var(--main-border)' }}>
                <td style={{ padding: '18px 0', color: 'var(--text-muted)', fontWeight: 600 }}>Dérive Calendaire</td>
                <td style={{ textAlign: 'right', fontWeight: 900, color: '#ef4444' }}>-33% (Retard)</td>
              </tr>
              <tr style={{ borderBottom: '1px solid var(--main-border)' }}>
                <td style={{ padding: '18px 0', color: 'var(--text-muted)', fontWeight: 600 }}>Tâches Restantes</td>
                <td style={{ textAlign: 'right', fontWeight: 900, color: 'var(--main-text)' }}>33 en attente</td>
              </tr>
              <tr style={{ borderBottom: '1px solid var(--main-border)' }}>
                <td style={{ padding: '18px 0', color: 'var(--text-muted)', fontWeight: 600 }}>Efficience des Coûts</td>
                <td style={{ textAlign: 'right', fontWeight: 900, color: '#f59e0b' }}>+26% Overconsumption</td>
              </tr>
            </tbody>
          </table>
        </div>

        <div style={{ background: 'var(--main-surface)', padding: '2.5rem', borderRadius: '20px', border: '1px solid var(--main-border)', boxShadow: 'var(--card-shadow)' }}>
          <h3 style={{ color: '#3b82f6', marginBottom: '2rem', fontSize: '1rem', fontWeight: 900, textTransform: 'uppercase' }}>🏗️ Progression par Phase</h3>
          <div style={{ display: 'flex', flexDirection: 'column', gap: '1.5rem' }}>
            {[{ p: 'Initialisation', v: 76 }, { p: 'Conception', v: 60 }, { p: 'Exécution', v: 31 }, { p: 'Contrôle', v: 32 }].map(i => (
              <div key={i.p}>
                <div style={{ display: 'flex', justifyContent: 'space-between', fontSize: '0.85rem', marginBottom: '8px' }}>
                  <span style={{ fontWeight: 700, color: 'var(--main-text)' }}>{i.p}</span>
                  <span style={{ fontWeight: 900, color: i.v > 70 ? '#10b981' : i.v > 50 ? '#3b82f6' : '#f59e0b' }}>{i.v}%</span>
                </div>
                <div style={{ height: '10px', background: 'var(--main-bg)', borderRadius: '5px', border: '1px solid var(--main-border)', overflow: 'hidden' }}>
                  <div style={{ width: `${i.v}%`, height: '100%', background: i.v > 70 ? '#10b981' : i.v > 50 ? '#3b82f6' : '#f59e0b' }}></div>
                </div>
              </div>
            ))}
          </div>
        </div>
      </div>
    </div>
  );
};

// ─── RapportFlashView ──────────────────────────────────────────
const RapportFlashView = ({ projectId, fileId, currentDossierId }: { projectId: string; fileId: string; currentDossierId: string | null }) => {
  const metadata = { projectName: 'SAP30 / CMRL', manager: 'Mme Dupont', version: 'V2.3', lastUpdate: '28/11/2022' };
  const dossierSubKey = currentDossierId ? `-${currentDossierId}` : "";
  const storageKey = `faab-rapport-flash-${projectId}-${fileId}${dossierSubKey}`;

  const defaultData = [
    { id: 1, action: 'Finaliser le paramétrage Finances (Lot A)', resp: 'Mr. Anselme (CFO)', ech: '23/07/2022', status: 'COMPLÉTÉ' },
    { id: 2, action: 'Validation du lot 1 Technique (Infrastructure)', resp: 'Mme. Durand (DSI)', ech: '15/08/2022', status: 'EN COURS' }
  ];

  const [data, setData] = useState(() => {
    const saved = localStorage.getItem(storageKey);
    return saved ? JSON.parse(saved) : defaultData;
  });

  useEffect(() => {
    localStorage.setItem(storageKey, JSON.stringify(data));
  }, [data, storageKey]);
  return (
    <div style={{ flex: 1, display: 'flex', flexDirection: 'column', background: 'var(--main-bg)', color: 'var(--main-text)' }}>
      <ExecutionHeader title="30 — Rapport Flash Hebdomadaire" metadata={metadata} />
      <div style={{ padding: '2.5rem', display: 'flex', gap: '2rem' }}>
        <div style={{ background: 'var(--main-surface)', padding: '2rem', borderRadius: '20px', border: '1px solid var(--main-border)', flex: 1, boxShadow: 'var(--card-shadow)', textAlign: 'center' }}>
          <h3 style={{ fontSize: '0.8rem', color: '#10b981', marginBottom: '1.5rem', fontWeight: 900, textTransform: 'uppercase', letterSpacing: '0.1em' }}>Santé Hebdomadaire du Projet</h3>
          <div style={{ display: 'flex', justifyContent: 'space-around', alignItems: 'center' }}>
            <div style={{ padding: '15px' }}>
              <div style={{ fontSize: '0.75rem', color: 'var(--text-muted)', fontWeight: 800, marginBottom: '8px' }}>BUDGET</div>
              <div style={{ fontSize: '2.5rem' }}>☀️</div>
              <div style={{ fontSize: '0.7rem', color: '#10b981', fontWeight: 800, marginTop: '5px' }}>OPTIMAL</div>
            </div>
            <div style={{ padding: '15px' }}>
              <div style={{ fontSize: '0.75rem', color: 'var(--text-muted)', fontWeight: 800, marginBottom: '8px' }}>PLANNING</div>
              <div style={{ fontSize: '2.5rem' }}>🌥️</div>
              <div style={{ fontSize: '0.7rem', color: '#f59e0b', fontWeight: 800, marginTop: '5px' }}>VIGILANCE</div>
            </div>
            <div style={{ padding: '15px' }}>
              <div style={{ fontSize: '0.75rem', color: 'var(--text-muted)', fontWeight: 800, marginBottom: '8px' }}>PÉRIMÈTRE</div>
              <div style={{ fontSize: '2.5rem' }}>☀️</div>
              <div style={{ fontSize: '0.7rem', color: '#10b981', fontWeight: 800, marginTop: '5px' }}>CONFORME</div>
            </div>
          </div>
        </div>
      </div>
      <div style={{ flex: 1, overflow: 'auto', padding: '0 2.5rem 2.5rem' }}>
        <div style={{ background: 'var(--main-surface)', borderRadius: '20px', overflow: 'hidden', border: '1px solid var(--main-border)', boxShadow: 'var(--card-shadow)', position: 'relative' }}>
          <div style={{ background: 'var(--excel-color)', padding: '18px 25px', fontSize: '0.85rem', fontWeight: 900, color: '#fff', display: 'flex', justifyContent: 'space-between', alignItems: 'center', textTransform: 'uppercase' }}>
            <span>Décisions et Réalisations Clés</span>
            <span style={{ fontSize: '0.75rem', opacity: 0.9 }}>Semaine 42 • Rapport Consolidé</span>
          </div>
          <table className="editable-table" style={{ width: '100%', borderCollapse: 'separate', borderSpacing: 0, fontSize: '0.9rem' }}>
            <thead>
              <tr style={{ background: 'rgba(0,0,0,0.03)', color: 'var(--text-muted)' }}>
                <th style={{ padding: '15px 25px', textAlign: 'left', fontWeight: 800, fontSize: '0.75rem', textTransform: 'uppercase' }}>Actions / Décisions Stratégiques</th>
                <th style={{ padding: '15px 25px', textAlign: 'left', fontWeight: 800, fontSize: '0.75rem', textTransform: 'uppercase' }}>Acteur Responsable</th>
                <th style={{ padding: '15px 25px', textAlign: 'center', fontWeight: 800, fontSize: '0.75rem', textTransform: 'uppercase' }}>Échéance Prévue</th>
                <th style={{ padding: '15px 25px', textAlign: 'center', fontWeight: 800, fontSize: '0.75rem', textTransform: 'uppercase' }}>Statut Actuel</th>
              </tr>
            </thead>
            <tbody style={{ color: 'var(--main-text)' }}>
              <tr style={{ borderBottom: '1px solid var(--main-border)' }}>
                <td style={{ padding: '20px 25px', fontWeight: 700 }}>Finaliser le paramétrage Finances (Lot A)</td>
                <td style={{ padding: '20px 25px', color: 'var(--text-muted)', fontWeight: 600 }}>Mr. Anselme (CFO)</td>
                <td style={{ padding: '20px 25px', textAlign: 'center', fontWeight: 800 }}>23/07/2022</td>
                <td style={{ padding: '20px 25px', textAlign: 'center' }}>
                  <span style={{ padding: '8px 16px', background: 'rgba(16, 185, 129, 0.12)', color: '#10b981', borderRadius: '30px', fontSize: '0.7rem', fontWeight: 950 }}>COMPLÉTÉ</span>
                </td>
              </tr>
              <tr style={{ borderBottom: '1px solid var(--main-border)', background: 'rgba(0,0,0,0.01)' }}>
                <td style={{ padding: '20px 25px', fontWeight: 700 }}>Validation du lot 1 Technique (Infrastructure)</td>
                <td style={{ padding: '20px 25px', color: 'var(--text-muted)', fontWeight: 600 }}>Comité de Pilotage</td>
                <td style={{ padding: '20px 25px', textAlign: 'center', fontWeight: 800 }}>28/07/2022</td>
                <td style={{ padding: '20px 25px', textAlign: 'center' }}>
                  <span style={{ padding: '8px 16px', background: 'rgba(59, 130, 246, 0.12)', color: '#3b82f6', borderRadius: '30px', fontSize: '0.7rem', fontWeight: 950 }}>EN COURS</span>
                </td>
              </tr>
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
};

// ─── BilanProjetView ──────────────────────────────────────────
const BilanProjetView = ({ page = 'option1', projectId, fileId, currentDossierId }: { page?: string, projectId: string, fileId: string, currentDossierId: string | null }) => {
  const metadata = { projectName: 'SAP30 / CMRL', manager: 'Chantal Dupont', version: 'V2.3', lastUpdate: '28/01/2022' };
  const dossierSubKey = currentDossierId ? `-${currentDossierId}` : "";
  const storageKey = `faab-bilan-projet-${projectId}-${fileId}${dossierSubKey}`;

  if (page === 'option2') {
    return (
      <div style={{ flex: 1, overflow: 'auto', background: 'var(--main-bg)', color: 'var(--main-text)' }}>
        <ExecutionHeader title="34 — Bilan de projet (REX / Satisfaction)" metadata={metadata} />
        <div style={{ padding: '1.25rem 2.5rem', background: 'var(--main-surface)', borderBottom: '1px solid var(--main-border)', display: 'flex', gap: '10px' }}>
          <div style={{ display: 'inline-flex', alignItems: 'center', gap: '8px', background: 'rgba(245, 158, 11, 0.12)', color: '#f59e0b', padding: '8px 18px', borderRadius: '30px', fontSize: '0.75rem', fontWeight: 900, textTransform: 'uppercase', letterSpacing: '0.05em' }}>
            📊 VUE ACTIVE : REX / SATISFACTION
          </div>
        </div>
        <div style={{ padding: '2.5rem', display: 'flex', flexDirection: 'column', gap: '2.5rem' }}>
          <section style={{ background: 'var(--main-surface)', padding: '2.5rem', borderRadius: '20px', border: '1px solid var(--main-border)', boxShadow: 'var(--card-shadow)' }}>
            <h3 style={{ color: 'var(--excel-color)', fontSize: '1.1rem', marginBottom: '2rem', borderBottom: '2px solid var(--main-border)', paddingBottom: '1rem', fontWeight: 900, textTransform: 'uppercase' }}>2. RETOUR D'EXPÉRIENCE (REX)</h3>
            <div style={{ display: 'flex', flexDirection: 'column', gap: '1.5rem' }}>
              <div style={{ padding: '1.5rem', background: 'rgba(16, 185, 129, 0.05)', borderRadius: '15px', borderLeft: '6px solid #10b981', boxShadow: '0 4px 6px rgba(0,0,0,0.02)' }}>
                <div style={{ color: '#10b981', fontWeight: 900, fontSize: '0.8rem', textTransform: 'uppercase', marginBottom: '8px' }}>✅ Points de Succès</div>
                <div style={{ fontSize: '0.95rem', lineHeight: '1.7', fontWeight: 500 }}>Adhésion forte des utilisateurs de la comptabilité dès la phase de recette. La formation a été un facteur clé de réussite.</div>
              </div>
              <div style={{ padding: '1.5rem', background: 'rgba(239, 68, 68, 0.05)', borderRadius: '15px', borderLeft: '6px solid #ef4444', boxShadow: '0 4px 6px rgba(0,0,0,0.02)' }}>
                <div style={{ color: '#ef4444', fontWeight: 900, fontSize: '0.8rem', textTransform: 'uppercase', marginBottom: '8px' }}>⚠️ Difficultés Rencontrées</div>
                <div style={{ fontSize: '0.95rem', lineHeight: '1.7', fontWeight: 500 }}>Délais de livraison des serveurs plus longs que prévu, impactant la phase de tests d'intégration de 2 semaines.</div>
              </div>
            </div>
          </section>
          <section style={{ background: 'var(--main-surface)', padding: '2.5rem', borderRadius: '20px', border: '1px solid var(--main-border)', boxShadow: 'var(--card-shadow)', textAlign: 'center' }}>
            <h3 style={{ color: '#f59e0b', fontSize: '1.1rem', marginBottom: '2rem', borderBottom: '2px solid var(--main-border)', paddingBottom: '1rem', fontWeight: 900, textTransform: 'uppercase' }}>3. SATISFACTION CLIENT</h3>
            <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'center', gap: '1rem' }}>
              <div style={{ fontSize: '5rem', fontWeight: 950, color: '#f59e0b', letterSpacing: '-0.02em' }}>3.8 <span style={{ fontSize: '2rem', opacity: 0.5 }}>/ 5</span></div>
              <div style={{ fontSize: '1.1rem', color: 'var(--text-muted)', fontWeight: 600, maxWidth: '500px' }}>Score moyen de satisfaction consolidé auprès des 12 parties prenantes clés.</div>
              <div style={{ marginTop: '1.5rem', height: '12px', width: '300px', background: 'var(--main-bg)', borderRadius: '6px', overflow: 'hidden', border: '1px solid var(--main-border)' }}>
                <div style={{ width: '76%', height: '100%', background: '#f59e0b', borderRadius: '6px' }}></div>
              </div>
            </div>
          </section>
        </div>
      </div>
    );
  }

  if (page === 'option3') {
    return (
      <div style={{ flex: 1, overflow: 'auto', background: 'var(--main-bg)', color: 'var(--main-text)' }}>
        <ExecutionHeader title="34 — Checklist de clôture" metadata={metadata} />
        <div style={{ padding: '1.25rem 2.5rem', background: 'var(--main-surface)', borderBottom: '1px solid var(--main-border)', display: 'flex', gap: '10px' }}>
          <div style={{ display: 'inline-flex', alignItems: 'center', gap: '8px', background: 'rgba(16, 185, 129, 0.12)', color: '#10b981', padding: '8px 18px', borderRadius: '30px', fontSize: '0.75rem', fontWeight: 900, textTransform: 'uppercase', letterSpacing: '0.05em' }}>
            📋 VUE ACTIVE : CHECKLIST DE CLÔTURE
          </div>
        </div>
        <div style={{ padding: '2.5rem' }}>
          <div style={{ background: 'var(--main-surface)', borderRadius: '20px', border: '1px solid var(--main-border)', overflow: 'hidden', boxShadow: 'var(--card-shadow)' }}>
            <table style={{ width: '100%', borderCollapse: 'separate', borderSpacing: 0, fontSize: '0.9rem' }}>
              <thead style={{ background: 'var(--excel-color)', color: '#fff' }}>
                <tr style={{ fontWeight: 800, textTransform: 'uppercase', fontSize: '0.75rem' }}>
                  <th style={{ padding: '18px 25px', textAlign: 'left' }}>Élément de contrôle / Jalons</th>
                  <th style={{ padding: '18px 25px', textAlign: 'center' }}>Vérifié</th>
                  <th style={{ padding: '18px 25px', textAlign: 'left' }}>Commentaires / Preuves</th>
                </tr>
              </thead>
              <tbody style={{ color: 'var(--main-text)' }}>
                {[
                  "Livrables signés par le client", "Rapport financier définitif", "PV de recette globale", "REX effectué avec l'équipe", "Libération des ressources", "Fermeture du code analytique"
                ].map((item, idx) => (
                  <tr key={idx} style={{ borderBottom: '1px solid var(--main-border)', background: idx % 2 === 0 ? 'transparent' : 'rgba(0,0,0,0.01)' }}>
                    <td style={{ padding: '18px 25px', fontWeight: 700 }}>{item}</td>
                    <td style={{ padding: '18px 25px', textAlign: 'center' }}>
                      <input type="checkbox" defaultChecked={idx < 4} style={{ width: '22px', height: '22px', accentColor: '#10b981', cursor: 'pointer' }} />
                    </td>
                    <td style={{ padding: '18px 25px', color: 'var(--text-muted)', fontStyle: 'italic', fontWeight: 500 }}>
                      {idx < 4 ? `Validé le ${20 - idx}/11/2022` : 'En attente de validation finale'}
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>
      </div>
    );
  }

  const generateArchive = () => {
    const archiveData = {
      projet: metadata.projectName,
      date: new Date().toISOString(),
      statut: 'Archivé Définitivement',
      cloture: 'Validée',
      rex: {
        pointsSucces: "Adhésion forte des utilisateurs dès la recette.",
        difficultes: "Délais de livraison des serveurs."
      },
      satisfaction: 3.8,
      timestamp: Date.now()
    };
    const blob = new Blob([JSON.stringify(archiveData, null, 2)], { type: 'application/json' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `archive_projet_${projectId}_${Date.now()}.json`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
    alert('✅ Archive consolidée du projet générée et téléchargée avec succès !');
  };

  if (page === 'option4') {
    return (
      <div style={{ flex: 1, overflow: 'auto', background: 'var(--main-bg)', color: 'var(--main-text)' }}>
        <ExecutionHeader title="34 — Archivage du projet" metadata={metadata} />
        <div style={{ padding: '1.25rem 2.5rem', background: 'var(--main-surface)', borderBottom: '1px solid var(--main-border)', display: 'flex', gap: '10px' }}>
          <div style={{ display: 'inline-flex', alignItems: 'center', gap: '8px', background: 'rgba(100, 116, 139, 0.12)', color: '#64748b', padding: '8px 18px', borderRadius: '30px', fontSize: '0.75rem', fontWeight: 900, textTransform: 'uppercase', letterSpacing: '0.05em' }}>
            📂 VUE ACTIVE : ARCHIVAGE DÉFINITIF
          </div>
        </div>
        <div style={{ padding: '5rem 2.5rem', display: 'flex', justifyContent: 'center' }}>
          <div style={{ maxWidth: '700px', width: '100%', background: 'var(--main-surface)', padding: '4rem', borderRadius: '32px', border: '2px dashed var(--main-border)', textAlign: 'center', boxShadow: 'var(--card-shadow)' }}>
            <div style={{ fontSize: '5rem', marginBottom: '2rem', filter: 'drop-shadow(0 10px 15px rgba(0,0,0,0.1))' }}>📦</div>
            <h2 style={{ fontSize: '2.2rem', fontWeight: 950, color: 'var(--main-text)', marginBottom: '1.5rem', letterSpacing: '-0.02em' }}>Générer l'Archive Consolidée</h2>
            <p style={{ color: 'var(--text-muted)', fontSize: '1.1rem', marginBottom: '3rem', fontWeight: 500, lineHeight: '1.6' }}>
              Tous les documents, livrables, rapports financiers et comptes-rendus seront compressés en un fichier unique (.zip) pour l'archivage définitif et légal du projet.
            </p>
            <button
              className="btn-action"
              onClick={generateArchive}
              style={{ padding: '18px 45px', background: 'linear-gradient(135deg, #10b981, #059669)', color: '#fff', border: 'none', borderRadius: '40px', fontWeight: 900, fontSize: '1rem', cursor: 'pointer', boxShadow: '0 10px 20px -5px rgba(16, 185, 129, 0.4)', transition: 'transform 0.2s cubic-bezier(0.175, 0.885, 0.32, 1.275)' }}
            >
              PRODUIRE L'ARCHIVE (.ZIP)
            </button>
          </div>
        </div>
      </div>
    );
  }

  return (
    <div style={{ flex: 1, overflow: 'auto', background: 'var(--main-bg)', color: 'var(--main-text)' }}>
      <ExecutionHeader title="34 — Bilan de projet (Analyse des écarts)" metadata={metadata} />
      <div style={{ padding: '1.25rem 2.5rem', background: 'var(--main-surface)', borderBottom: '1px solid var(--main-border)', display: 'flex', gap: '10px' }}>
        <div style={{ display: 'inline-flex', alignItems: 'center', gap: '8px', background: 'rgba(59, 130, 246, 0.12)', color: '#3b82f6', padding: '8px 18px', borderRadius: '30px', fontSize: '0.75rem', fontWeight: 900, textTransform: 'uppercase', letterSpacing: '0.05em' }}>
          🏁 VUE ACTIVE : ANALYSE DES ÉCARTS
        </div>
      </div>
      <div style={{ padding: '2.5rem' }}>
        <div style={{ background: 'var(--main-surface)', padding: '3rem', borderRadius: '24px', border: '1px solid var(--main-border)', boxShadow: 'var(--card-shadow)' }}>
          <h3 style={{ color: '#10b981', fontSize: '1.1rem', marginBottom: '2.5rem', borderBottom: '2px solid var(--main-border)', paddingBottom: '1rem', fontWeight: 950, textTransform: 'uppercase' }}>1. ANALYSE CRITIQUE DES ÉCARTS</h3>
          <div style={{ display: 'flex', flexDirection: 'column', gap: '1.5rem', fontSize: '1rem' }}>
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', padding: '20px 25px', background: 'var(--main-bg)', borderRadius: '16px', border: '1px solid var(--main-border)', transition: 'all 0.2s ease' }}>
              <span style={{ fontWeight: 700 }}>Les livrables ont-ils été livrés tels que prévus initialement ?</span>
              <span style={{ color: '#ef4444', fontWeight: 950, background: 'rgba(239, 68, 68, 0.12)', padding: '6px 20px', borderRadius: '30px', fontSize: '0.75rem', letterSpacing: '0.05em' }}>NON CONFORME</span>
            </div>
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', padding: '20px 25px', background: 'var(--main-bg)', borderRadius: '16px', border: '1px solid var(--main-border)', transition: 'all 0.2s ease' }}>
              <span style={{ fontWeight: 700 }}>Les objectifs stratégiques du projet ont-ils été pleinement atteints ?</span>
              <span style={{ color: '#10b981', fontWeight: 950, background: 'rgba(16, 185, 129, 0.12)', padding: '6px 20px', borderRadius: '30px', fontSize: '0.75rem', letterSpacing: '0.05em' }}>OUI - SUCCÈS</span>
            </div>
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', padding: '20px 25px', background: 'var(--main-bg)', borderRadius: '16px', border: '1px solid var(--main-border)', transition: 'all 0.2s ease' }}>
              <span style={{ fontWeight: 700 }}>Respect des délais globaux et des jalons critiques (Planning) ?</span>
              <span style={{ color: '#f59e0b', fontWeight: 950, background: 'rgba(245, 158, 11, 0.12)', padding: '6px 20px', borderRadius: '30px', fontSize: '0.75rem', letterSpacing: '0.05em' }}>ATTÉNUÉ / MITIGÉ</span>
            </div>
          </div>
        </div>
      </div>
    </div>
  );
};


function App() {
  const [activeTabId, setActiveTabId] = useState<string>('selection');
  const [selectedFile, setSelectedFile] = useState<ProjectFile | null>(null);
  const [activeOptionPage, setActiveOptionPage] = useState<string | null>(null);
  const [isSidebarOpen, setIsSidebarOpen] = useState(true);
  const [theme, setTheme] = useState<'white' | 'gray' | 'black' | 'red' | 'default'>(() => {
    return (localStorage.getItem('faab-theme') as any) || 'default';
  });

  useEffect(() => {
    localStorage.setItem('faab-theme', theme);
  }, [theme]);
  const [fileDossiers, setFileDossiers] = useState<Record<string, string | null>>(() => {
    const saved = localStorage.getItem('faab-file-dossiers');
    if (saved) return JSON.parse(saved);
    // Legacy support for existing keys
    return {
      'selection-projets': localStorage.getItem('faab-current-selection-dossier'),
      'matrice-faisabilite': localStorage.getItem('faab-current-faisabilite-dossier'),
      'prenantes-1': localStorage.getItem('faab-current-prenantes1-dossier'),
      'prenantes-2': localStorage.getItem('faab-current-prenantes2-dossier'),
      'risques-1': localStorage.getItem('faab-current-risques1-dossier'),
      'risques-2': localStorage.getItem('faab-current-risques2-dossier'),
    };
  });

  const selectDossierForFile = (fileId: string, dossierId: string, autoNavigate: boolean = false) => {
    const newDossiers = { ...fileDossiers, [fileId]: dossierId };
    setFileDossiers(newDossiers);
    localStorage.setItem('faab-file-dossiers', JSON.stringify(newDossiers));

    // Set legacy keys for compatibility if needed
    if (fileId === 'selection-projets') localStorage.setItem('faab-current-selection-dossier', dossierId);
    if (fileId === 'matrice-faisabilite') localStorage.setItem('faab-current-faisabilite-dossier', dossierId);
    if (fileId === 'prenantes-1') localStorage.setItem('faab-current-prenantes1-dossier', dossierId);
    if (fileId === 'prenantes-2') localStorage.setItem('faab-current-prenantes2-dossier', dossierId);
    if (fileId === 'risques-1') localStorage.setItem('faab-current-risques1-dossier', dossierId);
    if (fileId === 'risques-2') localStorage.setItem('faab-current-risques2-dossier', dossierId);

    if (autoNavigate) {
      setActiveOptionPage('option1');
    }
  };

  const [portfolios, setPortfolios] = useState<Portfolio[]>(() => {
    const saved = localStorage.getItem('faab-portfolios');
    if (saved) {
      try {
        const parsed = JSON.parse(saved);
        if (Array.isArray(parsed) && parsed.length > 0) return parsed;
      } catch (e) {
        console.error("Error parsing portfolios", e);
      }
    }
    return [{
      id: 'port-demo-1',
      name: 'Organisation FAAB',
      projects: [{ id: 'proj-demo-1', name: 'Projet FAAB' }]
    }];
  });


  const [activeProjectId, setActiveProjectIdState] = useState<string | null>(() => {
    const saved = localStorage.getItem('faab-active-project');
    if (saved) return saved;
    return 'proj-demo-1';
  });

  const setActiveProjectId = (id: string | null) => {
    setActiveProjectIdState(id);
    if (id) {
      localStorage.setItem('faab-active-project', id);
    } else {
      localStorage.removeItem('faab-active-project');
    }
  };

  // Auto-save portfolios
  useEffect(() => {
    localStorage.setItem('faab-portfolios', JSON.stringify(portfolios));
  }, [portfolios]);

  // Global selected project
  const activeProject = useMemo(() => {
    for (const port of portfolios) {
      const proj = port.projects.find(p => p.id === activeProjectId);
      if (proj) return proj;
    }
    return null;
  }, [portfolios, activeProjectId]);

  const activeTab = tabsData.find(t => t.id === activeTabId);

  const renderFileIcon = (type: string) => {
    if (type === 'word') return <WordIcon />;
    if (type === 'ppt') return <PptIcon />;
    return <ExcelIcon />;
  };

  return (
    <div className={`app-container theme-${theme} ${isSidebarOpen ? 'sidebar-open' : 'sidebar-collapsed'}`}>
      <div className="sidebar" style={{ transition: 'width 0.25s ease' }}>
        <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', padding: '0.8rem 1.2rem 1.2rem', borderBottom: '1px solid rgba(255,255,255,0.08)', marginBottom: '0.8rem' }}>
          <div style={{ display: 'flex', alignItems: 'center', gap: '10px', cursor: 'pointer' }} onClick={() => setSelectedFile(null)}>
            <img src="/logo.png" alt="FAAB PROJET" style={{ width: '38px', height: '38px', borderRadius: '8px', boxShadow: '0 4px 12px rgba(16,185,129,0.35)' }} />
            {isSidebarOpen && (
              <div style={{ display: 'flex', flexDirection: 'column' }}>
                <span style={{ fontWeight: 800, fontSize: '1.05rem', letterSpacing: '0.5px', background: 'linear-gradient(90deg, #10b981, #38bdf8)', WebkitBackgroundClip: 'text', WebkitTextFillColor: 'transparent' }}>
                  FAAB PROJET
                </span>
                <span style={{ fontSize: '0.65rem', color: '#94a3b8', textTransform: 'uppercase', letterSpacing: '1px', fontWeight: 600 }}>
                  Kit Chef de Projet
                </span>
              </div>
            )}
          </div>
          <button
            onClick={() => setIsSidebarOpen(!isSidebarOpen)}
            style={{ background: 'transparent', border: 'none', color: 'var(--text-muted)', cursor: 'pointer', fontSize: '1.1rem', padding: '4px' }}
            title={isSidebarOpen ? 'Réduire le menu' : 'Agrandir le menu'}
          >
            {isSidebarOpen ? '◀' : '▶'}
          </button>
        </div>

        {isSidebarOpen && (
          <div className="tab-list">
            {/* Bouton d'accueil Cockpit fixe */}
            <div
              className={`tab-item ${!selectedFile ? 'active' : ''}`}
              onClick={() => { setSelectedFile(null); }}
              style={{
                marginBottom: '0.8rem',
                background: !selectedFile ? 'linear-gradient(135deg, rgba(16, 185, 129, 0.25), rgba(5, 150, 105, 0.15))' : 'rgba(255,255,255,0.03)',
                borderColor: !selectedFile ? '#10b981' : 'transparent',
                fontWeight: 700,
                cursor: 'pointer'
              }}
            >
              <div className="tab-title" style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
                <span>🏠</span>
                <span>Cockpit & Vue d'ensemble</span>
              </div>
            </div>

            {tabsData.map(tab => (
              <div key={tab.id} className="tab-wrapper">
                <div
                  className={`tab-item ${activeTabId === tab.id ? 'active' : ''}`}
                  onClick={() => { setActiveTabId(tab.id); }}
                >
                  <div className="tab-title">{tab.title}</div>
                  {tab.files.length > 0 && (
                    <div className="tab-chevron">{activeTabId === tab.id ? '▼' : '▶'}</div>
                  )}
                </div>

                {activeTabId === tab.id && tab.files.length > 0 && (
                  <div className="sub-tab-list">
                    {tab.files.map(file => {
                      const isSelected = selectedFile?.id === file.id;
                      return (
                        <div key={file.id} style={{ display: 'flex', flexDirection: 'column' }}>
                          <div
                            className={`sub-tab-item ${isSelected ? 'active' : ''}`}
                            onClick={() => {
                              setSelectedFile(file);
                              setActiveOptionPage('option1');
                            }}
                            title={file.name}
                          >
                            <span className={`file-icon-small ${file.type}`} style={{ marginRight: '8px' }}>{renderFileIcon(file.type)}</span>
                            <span className="sub-tab-name">{file.name}</span>
                          </div>
                        </div>
                      );
                    })}
                  </div>
                )}
              </div>
            ))}
          </div>
        )}

        {!isSidebarOpen && (
          <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'center', gap: '1rem', marginTop: '1rem' }}>
            {tabsData.map(tab => (
              <div
                key={tab.id}
                onClick={() => { setIsSidebarOpen(true); setActiveTabId(tab.id); }}
                style={{ cursor: 'pointer', fontSize: '1.2rem', opacity: activeTabId === tab.id ? 1 : 0.5 }}
                title={tab.title}
              >
                📁
              </div>
            ))}
          </div>
        )}
      </div>

      <div className="main-content">
        <header className="header" style={{ justifyContent: 'space-between', alignItems: 'center', padding: '0.8rem 1.8rem', background: 'rgba(15, 23, 42, 0.75)', borderBottom: '1px solid var(--main-border)', backdropFilter: 'blur(10px)' }}>
          <div style={{ display: 'flex', alignItems: 'center', gap: '12px', flexWrap: 'wrap' }}>
            {selectedFile ? (
              <>
                <button
                  onClick={() => setSelectedFile(null)}
                  style={{
                    background: 'rgba(255, 255, 255, 0.08)',
                    border: '1px solid rgba(255, 255, 255, 0.15)',
                    color: '#f8fafc',
                    padding: '0.45rem 0.85rem',
                    borderRadius: '8px',
                    cursor: 'pointer',
                    fontSize: '0.82rem',
                    fontWeight: 600,
                    display: 'flex',
                    alignItems: 'center',
                    gap: '6px',
                    transition: 'background 0.15s ease'
                  }}
                  onMouseEnter={(e) => e.currentTarget.style.background = 'rgba(255, 255, 255, 0.15)'}
                  onMouseLeave={(e) => e.currentTarget.style.background = 'rgba(255, 255, 255, 0.08)'}
                >
                  ← Cockpit
                </button>
                <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
                  <span style={{ fontSize: '1.05rem', fontWeight: 700, color: '#f8fafc' }}>{selectedFile.name}</span>
                  <span style={{ background: 'rgba(16, 185, 129, 0.2)', color: '#10b981', padding: '2px 8px', borderRadius: '4px', fontSize: '0.72rem', fontWeight: 700 }}>
                    {activeTab?.title}
                  </span>
                </div>
              </>
            ) : (
              <div style={{ display: 'flex', alignItems: 'center', gap: '10px' }}>
                <span style={{ fontSize: '1.2rem' }}>📊</span>
                <span style={{ fontSize: '1.1rem', fontWeight: 700, color: '#f8fafc' }}>
                  Tableau de Bord & Cockpit Kit Chef de Projet
                </span>
              </div>
            )}
          </div>

          <div style={{ display: 'flex', alignItems: 'center', gap: '1rem' }}>
            {selectedFile && selectedFile.type !== 'word' && selectedFile.type !== 'ppt' && (
              <div style={{ display: 'flex', gap: '6px', background: 'rgba(0,0,0,0.25)', padding: '4px', borderRadius: '8px', border: '1px solid rgba(255,255,255,0.06)' }}>
                {getFileOptions(selectedFile.id).map(opt => {
                  const isActive = activeOptionPage === opt.id || (!activeOptionPage && opt.id === 'option1');
                  return (
                    <button
                      key={opt.id}
                      onClick={() => setActiveOptionPage(opt.id)}
                      style={{
                        background: isActive ? '#10b981' : 'transparent',
                        color: isActive ? '#ffffff' : '#94a3b8',
                        border: 'none',
                        padding: '0.4rem 0.8rem',
                        borderRadius: '6px',
                        fontSize: '0.8rem',
                        fontWeight: 600,
                        cursor: 'pointer',
                        transition: 'all 0.15s ease'
                      }}
                    >
                      {opt.label}
                    </button>
                  );
                })}
              </div>
            )}

            <div style={{ display: 'flex', gap: '0.4rem', background: 'rgba(255,255,255,0.05)', padding: '4px', borderRadius: '30px', border: '1px solid var(--main-border)' }}>
              {[
                { id: 'white', color: '#fff', label: 'Blanc' },
                { id: 'gray', color: '#cbd5e1', label: 'Gris' },
                { id: 'black', color: '#020617', label: 'Noir' },
                { id: 'red', color: '#7f1d1d', label: 'Rouge' },
                { id: 'default', color: '#6366f1', label: 'Défaut' }
              ].map(t => (
                <button
                  key={t.id}
                  onClick={() => setTheme(t.id as any)}
                  title={t.label}
                  style={{
                    width: '20px',
                    height: '20px',
                    borderRadius: '50%',
                    background: t.color,
                    border: theme === t.id ? '2px solid #fff' : '1px solid rgba(0,0,0,0.1)',
                    cursor: 'pointer',
                    padding: 0,
                    boxShadow: theme === t.id ? '0 0 8px rgba(0,0,0,0.3)' : 'none',
                    transition: 'all 0.15s'
                  }}
                />
              ))}
            </div>
          </div>
        </header>

        <div className="content-area" style={{ padding: selectedFile ? '0' : '2.5rem', display: 'flex', flexDirection: 'column' }}>
          {selectedFile ? (
            <div className="viewer-body" style={{ flex: 1, display: 'flex', flexDirection: 'column', height: '100%' }}>
              <div style={{ flex: 1, overflow: 'hidden', position: 'relative', display: 'flex', flexDirection: 'column', background: 'var(--main-bg)' }}>
                {activeOptionPage === 'option_dossiers' && (
                  <SelectionDossiersView
                    key={`${activeProjectId}-${selectedFile.id}`}
                    projectId={activeProjectId!}
                    fileId={selectedFile.id}
                    onSelectDossier={(id, autoNav = false) => selectDossierForFile(selectedFile.id, id, autoNav)}
                    onOpenDossier={(id) => selectDossierForFile(selectedFile.id, id, true)}
                    currentDossierId={fileDossiers[selectedFile.id] || null}
                    moduleType={(selectedFile.id === 'selection-projets' ? 'selection' : 'faisabilite') as any}
                    title={`Dossiers : ${selectedFile.name}`}
                  />
                )}
                {activeOptionPage !== 'option_dossiers' && (
                  (() => {
                    const fid = selectedFile.id;
                    const isSelection = fid === 'selection-projets';
                    const isFaisabilite = fid === 'matrice-faisabilite';
                    const isStakeholder1 = fid === 'prenantes-1';
                    const isStakeholder2 = fid === 'prenantes-2';
                    const isPlanning = fid === 'macro-planning';
                    const isGantt = fid === 'gantt';
                    const isRisques1 = fid === 'risques-1';
                    const isRisques2 = fid === 'risques-2';
                    const isPlanCom = fid === 'comm-fr' || fid === 'comm-act-en';
                    const isCompetences = fid === 'competences';
                    const isWBS = fid === 'wbs';
                    const isRACI = fid === 'raci';
                    const isPlanCharge = fid === 'plan-charge';
                    const isMatriceExig = fid === 'matrice-exigences';

                    const isLivrables = fid === 'suivi-livrables';
                    const isAMDEC = fid === 'amdec';
                    const isPlanTest = fid === 'plan-test';
                    const isChangements = fid === 'chg-req-reg';
                    const isActions = fid === 'suivi-actions';

                    const isBudget = fid === 'suivi-budget';
                    const isTB = fid === 'tb';
                    const isRapportFlash = fid === 'rapport-flash-xls';
                    const isBilan = fid === 'bilan';

                    const currentDossierId = fileDossiers[fid];

                    if (isSelection || isFaisabilite) {
                      let mappedMode: any = 'evaluation';
                      if (activeOptionPage === 'option_liste') mappedMode = 'database';
                      if (activeOptionPage === 'option_dashboard') mappedMode = 'dashboard';
                      if (activeOptionPage === 'option_dossiers') mappedMode = 'dossiers';
                      if (activeOptionPage === 'option_settings') mappedMode = 'settings';
                      if (activeOptionPage === 'option1') mappedMode = 'evaluation';
                      if (activeOptionPage === 'option2') mappedMode = 'criteria';
                      if (activeOptionPage === 'option3' && isSelection) mappedMode = 'settings';

                      if (isSelection) {
                        return <SelectionEvaluationView
                          fileId={fid}
                          projectId={activeProjectId!}
                          currentDossierId={currentDossierId}
                          initialView={mappedMode}
                          onDossierChange={(id) => selectDossierForFile(fid, id, false)}
                        />;
                      } else {
                        return <FaisabiliteMatriceView
                          fileId={fid}
                          projectId={activeProjectId!}
                          currentDossierId={currentDossierId}
                          initialView={mappedMode}
                          onDossierChange={(id) => selectDossierForFile(fid, id, false)}
                        />;
                      }
                    }

                    if (isStakeholder1 || isStakeholder2) {
                      let mappedMode: any = 'registry';
                      if (isStakeholder1) {
                        if (activeOptionPage === 'option1') mappedMode = 'registry';
                        if (activeOptionPage === 'option2') mappedMode = 'matrix';
                        if (activeOptionPage === 'option_dossiers') mappedMode = 'dossiers';
                        if (activeOptionPage === 'option3') mappedMode = 'settings';
                        return <PartiesPrenantesTableView
                          fileId={fid}
                          projectId={activeProjectId!}
                          currentDossierId={currentDossierId}
                          initialView={mappedMode}
                          onDossierChange={(id) => selectDossierForFile(fid, id, false)}
                        />;
                      } else {
                        if (activeOptionPage === 'option1') mappedMode = 'cartography';
                        if (activeOptionPage === 'option2') mappedMode = 'graph';
                        if (activeOptionPage === 'option_dossiers') mappedMode = 'dossiers';
                        return <CartographieDecideursTableView
                          fileId={fid}
                          projectId={activeProjectId!}
                          currentDossierId={currentDossierId}
                          initialView={mappedMode}
                          onDossierChange={(id) => selectDossierForFile(fid, id, false)}
                        />;
                      }
                    }

                    if (isPlanning || isGantt) return <MacroPlanningView fileId={fid} projectId={activeProjectId!} currentDossierId={currentDossierId} initialView={activeOptionPage || 'option1'} />;

                    if (isRisques1 || isRisques2) {
                      let mappedMode: any = 'registry';
                      if (isRisques1) {
                        if (activeOptionPage === 'option1') mappedMode = 'guide';
                        if (activeOptionPage === 'option2') mappedMode = 'scales';
                        if (activeOptionPage === 'option3') mappedMode = 'registry';
                        if (activeOptionPage === 'option4') mappedMode = 'problems';
                        if (activeOptionPage === 'option5') mappedMode = 'matrix';
                        if (activeOptionPage === 'option_dossiers') mappedMode = 'dossiers';
                        return <RegistreRisquesTableView
                          fileId={fid}
                          projectId={activeProjectId!}
                          currentDossierId={currentDossierId}
                          initialView={mappedMode}
                          onDossierChange={(id) => selectDossierForFile(fid, id, false)}
                          isVariante1
                        />;
                      } else {
                        if (activeOptionPage === 'option1') mappedMode = 'registry';
                        if (activeOptionPage === 'option2') mappedMode = 'dashboard';
                        if (activeOptionPage === 'option_dossiers') mappedMode = 'dossiers';
                        if (activeOptionPage === 'option3') mappedMode = 'settings';

                        return <RegistreRisquesTableView
                          fileId={fid}
                          projectId={activeProjectId!}
                          currentDossierId={currentDossierId}
                          initialView={mappedMode}
                          onDossierChange={(id) => selectDossierForFile(fid, id, false)}
                        />;
                      }
                    }
                    if (isPlanCom) return <CommunicationActionPlanView projectId={activeProjectId!} fileId={fid} currentDossierId={currentDossierId} page={activeOptionPage || 'option1'} />;
                    if (isCompetences) return <CompetencesMatriceView projectId={activeProjectId!} fileId={fid} currentDossierId={currentDossierId} page={activeOptionPage || 'option1'} />;
                    if (isWBS) return <WBSPBSView projectId={activeProjectId!} fileId={fid} currentDossierId={currentDossierId} page={activeOptionPage || 'option1'} />;
                    if (isRACI) return <RACIView projectId={activeProjectId!} fileId={fid} currentDossierId={currentDossierId} page={activeOptionPage || 'option1'} />;
                    if (isPlanCharge) return <PlanChargeView projectId={activeProjectId!} fileId={fid} currentDossierId={currentDossierId} page={activeOptionPage || 'option2'} onPageChange={(p) => setActiveOptionPage(p)} />;

                    if (isMatriceExig) return <MatriceExigencesView projectId={activeProjectId!} fileId={fid} currentDossierId={currentDossierId} page={activeOptionPage || 'option1'} />;

                    if (isLivrables) {
                      let p: any = 'livrables';
                      if (activeOptionPage === 'option2') p = 'parametrage';
                      return <SuiviLivrablesView projectId={activeProjectId!} fileId={fid} currentDossierId={currentDossierId} page={p} />;
                    }

                    if (isAMDEC) {
                      let p: any = 'echelles';
                      if (activeOptionPage === 'option1') p = 'echelles';
                      if (activeOptionPage === 'option2') p = 'logiciel';
                      if (activeOptionPage === 'option3') p = 'materiel';
                      if (activeOptionPage === 'option4') p = 'repartition';
                      return <AMDECView projectId={activeProjectId!} fileId={fid} currentDossierId={currentDossierId} page={p} />;
                    }

                    if (isPlanTest) {
                      let p: any = 'integration';
                      if (activeOptionPage === 'option2') p = 'consolidation';
                      return <PlanTestView projectId={activeProjectId!} fileId={fid} currentDossierId={currentDossierId} page={p} />;
                    }

                    if (isChangements) {
                      let p: any = 'demandes';
                      if (activeOptionPage === 'option2') p = 'parametrage';
                      return <ChangementView projectId={activeProjectId!} fileId={fid} currentDossierId={currentDossierId} page={p} />;
                    }

                    if (isActions) {
                      let p: any = 'actions';
                      if (activeOptionPage === 'option2') p = 'parametrage';
                      return <ActionsView projectId={activeProjectId!} fileId={fid} currentDossierId={currentDossierId} page={p} />;
                    }

                    if (isBudget) {
                      let sub = 'rapport';
                      if (activeOptionPage === 'option2') sub = 'cr';
                      else if (activeOptionPage === 'option3') sub = 'va';
                      else if (activeOptionPage === 'option4') sub = 'vp';
                      return <SuiviBudgetMatrix projectId={activeProjectId!} fileId={fid} currentDossierId={currentDossierId} activeSubTab={sub} />;
                    }

                    if (isTB) {
                      let sub = 'global';
                      if (activeOptionPage === 'option2') sub = 'deviations';
                      else if (activeOptionPage === 'option3') sub = 'costs';
                      else if (activeOptionPage === 'option4') sub = 'tasks';
                      return <TableauDeBordMatrix projectId={activeProjectId!} fileId={fid} currentDossierId={currentDossierId} activeSubTab={sub} />;
                    }

                    if (isRapportFlash) {
                      return <RapportFlashMatrix projectId={activeProjectId!} fileId={fid} currentDossierId={currentDossierId} activeSubTab="rapport" />;
                    }

                    if (isBilan) {
                      let sub = 'bilan';
                      if (activeOptionPage === 'option2') sub = 'rex';
                      else if (activeOptionPage === 'option3') sub = 'satisfaction';
                      else if (activeOptionPage === 'option4') sub = 'cloture';
                      return <BilanProjetMatrix projectId={activeProjectId!} fileId={fid} currentDossierId={currentDossierId} activeSubTab={sub} />;
                    }

                    if (selectedFile.type === 'word' || fid.startsWith('doc-')) {
                      return <ProjectReportGenerator projectId={activeProjectId!} fileId={fid} />;
                    }

                    if (selectedFile.type === 'ppt' || fid.startsWith('ppt-')) {
                      return <ProjectTrainingViewer initialCourseId={fid} />;
                    }

                    return <div style={{ padding: '2rem' }}>Contenu pour {fid}</div>;
                  })()
                )}
              </div>
            </div>
          ) : (
            <div style={{ flex: 1, overflowY: 'auto' }}>
              <Cockpit
                projectName={activeProject?.name || "Projet FAAB"}
                onSelectFile={(tabId, file) => {
                  setActiveTabId(tabId);
                  setSelectedFile(file);
                  setActiveOptionPage('option1');
                }}
              />
            </div>
          )}
        </div>
      </div>
    </div>
  );
}

export default App;
