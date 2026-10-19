import React from 'react';
import { tabsData, type ProjectFile } from './data';

interface CockpitProps {
  onSelectFile: (tabId: string, file: ProjectFile) => void;
  projectName?: string;
}

export const Cockpit: React.FC<CockpitProps> = ({ onSelectFile, projectName = "Projet FAAB" }) => {
  return (
    <div style={{ padding: '2rem 2.5rem', maxWidth: '1400px', margin: '0 auto', width: '100%' }}>
      {/* Hero Banner */}
      <div style={{
        background: 'linear-gradient(135deg, rgba(16, 185, 129, 0.15) 0%, rgba(15, 23, 42, 0.8) 100%)',
        border: '1px solid rgba(16, 185, 129, 0.3)',
        borderRadius: '16px',
        padding: '1.8rem 2rem',
        marginBottom: '2rem',
        display: 'flex',
        alignItems: 'center',
        justifyContent: 'space-between',
        backdropFilter: 'blur(10px)',
        boxShadow: '0 8px 32px rgba(0, 0, 0, 0.2)'
      }}>
        <div style={{ display: 'flex', alignItems: 'center', gap: '1.2rem' }}>
          <img src="/logo.png" alt="FAAB PROJET" style={{ width: '64px', height: '64px', borderRadius: '14px', boxShadow: '0 6px 20px rgba(16, 185, 129, 0.35)' }} />
          <div>
            <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
              <h1 style={{ margin: 0, fontSize: '1.8rem', fontWeight: 800, background: 'linear-gradient(90deg, #10b981, #38bdf8)', WebkitBackgroundClip: 'text', WebkitTextFillColor: 'transparent' }}>
                FAAB PROJET
              </h1>
              <span style={{ background: 'rgba(16, 185, 129, 0.2)', color: '#10b981', padding: '3px 10px', borderRadius: '20px', fontSize: '0.75rem', fontWeight: 700, border: '1px solid rgba(16, 185, 129, 0.4)' }}>
                KIT CHEF PROJET v3.0
              </span>
            </div>
            <p style={{ margin: '6px 0 0', color: '#94a3b8', fontSize: '0.95rem' }}>
              Plateforme de pilotage haute performance intégrée • Projet actif : <strong style={{ color: '#f8fafc' }}>{projectName}</strong>
            </p>
          </div>
        </div>

        <div style={{ display: 'flex', gap: '0.8rem', flexWrap: 'wrap' }}>
          <button
            onClick={() => {
              const execTab = tabsData.find(t => t.id === 'execution');
              const livrableFile = execTab?.files.find(f => f.id === 'suivi-livrables');
              if (execTab && livrableFile) onSelectFile(execTab.id, livrableFile);
            }}
            style={{
              background: 'linear-gradient(135deg, #10b981, #059669)',
              color: '#fff',
              border: 'none',
              padding: '0.6rem 1.1rem',
              borderRadius: '8px',
              fontWeight: 600,
              fontSize: '0.82rem',
              cursor: 'pointer',
              display: 'flex',
              alignItems: 'center',
              gap: '6px',
              boxShadow: '0 4px 14px rgba(16, 185, 129, 0.35)',
              transition: 'transform 0.15s ease'
            }}
            onMouseEnter={(e) => e.currentTarget.style.transform = 'translateY(-2px)'}
            onMouseLeave={(e) => e.currentTarget.style.transform = 'translateY(0)'}
          >
            📋 Suivi des Livrables
          </button>

          <button
            onClick={() => {
              const initTab = tabsData.find(t => t.id === 'initialisation');
              const charteFile = initTab?.files.find(f => f.id === 'doc-charte-projet');
              if (initTab && charteFile) onSelectFile(initTab.id, charteFile);
            }}
            style={{
              background: 'linear-gradient(135deg, #0284c7, #2563eb)',
              color: '#fff',
              border: 'none',
              padding: '0.6rem 1.1rem',
              borderRadius: '8px',
              fontWeight: 600,
              fontSize: '0.82rem',
              cursor: 'pointer',
              display: 'flex',
              alignItems: 'center',
              gap: '6px',
              boxShadow: '0 4px 14px rgba(2, 132, 199, 0.35)',
              transition: 'transform 0.15s ease'
            }}
            onMouseEnter={(e) => e.currentTarget.style.transform = 'translateY(-2px)'}
            onMouseLeave={(e) => e.currentTarget.style.transform = 'translateY(0)'}
          >
            📄 Rapports & PDF
          </button>

          <button
            onClick={() => {
              const initTab = tabsData.find(t => t.id === 'initialisation');
              const pptFile = initTab?.files.find(f => f.id === 'ppt-lancement');
              if (initTab && pptFile) onSelectFile(initTab.id, pptFile);
            }}
            style={{
              background: 'linear-gradient(135deg, #f59e0b, #d97706)',
              color: '#fff',
              border: 'none',
              padding: '0.6rem 1.1rem',
              borderRadius: '8px',
              fontWeight: 600,
              fontSize: '0.82rem',
              cursor: 'pointer',
              display: 'flex',
              alignItems: 'center',
              gap: '6px',
              boxShadow: '0 4px 14px rgba(245, 158, 11, 0.35)',
              transition: 'transform 0.15s ease'
            }}
            onMouseEnter={(e) => e.currentTarget.style.transform = 'translateY(-2px)'}
            onMouseLeave={(e) => e.currentTarget.style.transform = 'translateY(0)'}
          >
            🎓 Formations PPT
          </button>

          <button
            onClick={() => {
              const planTab = tabsData.find(t => t.id === 'planification');
              const ganttFile = planTab?.files.find(f => f.id === 'gantt');
              if (planTab && ganttFile) onSelectFile(planTab.id, ganttFile);
            }}
            style={{
              background: 'rgba(255, 255, 255, 0.08)',
              color: '#f8fafc',
              border: '1px solid rgba(255, 255, 255, 0.15)',
              padding: '0.6rem 1.1rem',
              borderRadius: '8px',
              fontWeight: 600,
              fontSize: '0.82rem',
              cursor: 'pointer',
              display: 'flex',
              alignItems: 'center',
              gap: '6px',
              transition: 'background 0.15s ease'
            }}
            onMouseEnter={(e) => e.currentTarget.style.background = 'rgba(255, 255, 255, 0.15)'}
            onMouseLeave={(e) => e.currentTarget.style.background = 'rgba(255, 255, 255, 0.08)'}
          >
            📊 Planning Gantt
          </button>
        </div>
      </div>

      {/* KPI Cards Row */}
      <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(240px, 1fr))', gap: '1rem', marginBottom: '2.5rem' }}>
        {[
          { title: "Livrables du Projet", val: "43", sub: "28 livrés • 12 en cours • 3 retard", color: "#10b981", badge: "65% Terminé", icon: "📋" },
          { title: "Actions Opérationnelles", val: "8", sub: "2 terminées • 4 en cours", color: "#38bdf8", badge: "Priorité Haute", icon: "✅" },
          { title: "Registre des Risques", val: "18", sub: "2 critiques • 5 modérés", color: "#f59e0b", badge: "Sous contrôle", icon: "🛡️" },
          { title: "Budget & Conformité", val: "94%", sub: "Consommation maîtrisée", color: "#a855f7", badge: "Conforme", icon: "💰" }
        ].map((card, i) => (
          <div key={i} style={{
            background: 'linear-gradient(135deg, rgba(30, 41, 59, 0.8) 0%, rgba(15, 23, 42, 0.9) 100%)',
            border: '1px solid rgba(255, 255, 255, 0.08)',
            borderRadius: '12px',
            padding: '1.2rem',
            position: 'relative',
            overflow: 'hidden'
          }}>
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start', marginBottom: '0.5rem' }}>
              <span style={{ fontSize: '0.8rem', color: '#94a3b8', fontWeight: 600, textTransform: 'uppercase' }}>{card.title}</span>
              <span style={{ fontSize: '1.2rem' }}>{card.icon}</span>
            </div>
            <div style={{ fontSize: '1.8rem', fontWeight: 800, color: card.color, marginBottom: '0.3rem' }}>
              {card.val}
            </div>
            <div style={{ fontSize: '0.78rem', color: '#cbd5e1' }}>
              {card.sub}
            </div>
            <div style={{
              position: 'absolute',
              top: '12px',
              right: '38px',
              background: 'rgba(255,255,255,0.06)',
              padding: '2px 8px',
              borderRadius: '10px',
              fontSize: '0.65rem',
              color: card.color,
              fontWeight: 700
            }}>
              {card.badge}
            </div>
          </div>
        ))}
      </div>

      {/* 6 Phases Workflow Matrix (Direct from KIT CHEF PROJET) */}
      <div style={{ marginBottom: '1.5rem', display: 'flex', alignItems: 'center', justifyContent: 'space-between' }}>
        <div>
          <h2 style={{ fontSize: '1.3rem', fontWeight: 700, margin: 0, color: '#f8fafc' }}>
            Cycle de Vie du Projet • 6 Phases KIT CHEF PROJET
          </h2>
          <p style={{ margin: '4px 0 0', color: '#94a3b8', fontSize: '0.85rem' }}>
            Cliquez sur n'importe quel classeur pour ouvrir instantanément l'outil et ses données.
          </p>
        </div>
      </div>

      <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(360px, 1fr))', gap: '1.2rem' }}>
        {tabsData.map((tab, idx) => {
          const colors = ['#10b981', '#38bdf8', '#818cf8', '#f59e0b', '#ec4899', '#14b8a6'];
          const phaseColor = colors[idx % colors.length];

          return (
            <div key={tab.id} style={{
              background: 'rgba(30, 41, 59, 0.6)',
              border: `1px solid rgba(255, 255, 255, 0.08)`,
              borderTop: `4px solid ${phaseColor}`,
              borderRadius: '12px',
              padding: '1.2rem',
              display: 'flex',
              flexDirection: 'column',
              boxShadow: '0 4px 16px rgba(0,0,0,0.1)'
            }}>
              <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginBottom: '0.8rem' }}>
                <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
                  <span style={{
                    background: `${phaseColor}20`,
                    color: phaseColor,
                    fontWeight: 800,
                    fontSize: '0.72rem',
                    padding: '2px 8px',
                    borderRadius: '6px'
                  }}>
                    PHASE 0{idx + 1}
                  </span>
                  <h3 style={{ margin: 0, fontSize: '1.05rem', fontWeight: 700, color: '#f8fafc' }}>
                    {tab.title}
                  </h3>
                </div>
                <span style={{ fontSize: '0.75rem', color: '#64748b', fontWeight: 600 }}>
                  {tab.files.length} outils
                </span>
              </div>

              <div style={{ display: 'flex', flexDirection: 'column', gap: '0.4rem', flex: 1 }}>
                {tab.files.map(file => (
                  <button
                    key={file.id}
                    onClick={() => onSelectFile(tab.id, file)}
                    style={{
                      background: 'rgba(255, 255, 255, 0.04)',
                      border: '1px solid rgba(255, 255, 255, 0.06)',
                      borderRadius: '8px',
                      padding: '0.6rem 0.8rem',
                      display: 'flex',
                      alignItems: 'center',
                      justifyContent: 'space-between',
                      cursor: 'pointer',
                      textAlign: 'left',
                      color: '#f1f5f9',
                      fontSize: '0.84rem',
                      fontWeight: 500,
                      transition: 'all 0.15s ease'
                    }}
                    onMouseEnter={(e) => {
                      e.currentTarget.style.background = 'rgba(255, 255, 255, 0.08)';
                      e.currentTarget.style.borderColor = phaseColor;
                      e.currentTarget.style.transform = 'translateX(4px)';
                    }}
                    onMouseLeave={(e) => {
                      e.currentTarget.style.background = 'rgba(255, 255, 255, 0.04)';
                      e.currentTarget.style.borderColor = 'rgba(255, 255, 255, 0.06)';
                      e.currentTarget.style.transform = 'translateX(0)';
                    }}
                  >
                    <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
                      <span style={{ fontSize: '0.95rem' }}>
                        {file.type === 'word' ? '📄' : file.type === 'ppt' ? '🎓' : '📗'}
                      </span>
                      <span>{file.name}</span>
                    </div>
                    <div style={{ display: 'flex', alignItems: 'center', gap: '6px' }}>
                      {file.type === 'word' && (
                        <span style={{ fontSize: '0.62rem', background: 'rgba(56, 189, 248, 0.15)', color: '#38bdf8', border: '1px solid rgba(56, 189, 248, 0.3)', padding: '1px 5px', borderRadius: '4px', fontWeight: 700 }}>
                          PDF
                        </span>
                      )}
                      {file.type === 'ppt' && (
                        <span style={{ fontSize: '0.62rem', background: 'rgba(245, 158, 11, 0.15)', color: '#f59e0b', border: '1px solid rgba(245, 158, 11, 0.3)', padding: '1px 5px', borderRadius: '4px', fontWeight: 700 }}>
                          FORMATION
                        </span>
                      )}
                      <span style={{ color: '#64748b', fontSize: '0.8rem' }}>→</span>
                    </div>
                  </button>
                ))}
              </div>
            </div>
          );
        })}
      </div>
    </div>
  );
};
