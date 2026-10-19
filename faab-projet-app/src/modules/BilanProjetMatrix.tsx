import React, { useState, useEffect, useMemo } from 'react';

interface BilanProjetMatrixProps {
  projectId: string;
  fileId: string;
  currentDossierId: string | null;
  activeSubTab?: string;
}

interface EcartItem {
  id: string;
  question: string;
  reponse: 'Oui' | 'Non';
  explication: string;
}

interface RiskReviewItem {
  id: number;
  risque: string;
  impactPlanifie: string;
  impactReel: string;
}

interface RexItem {
  domaine: string;
  lecons: string;
  recommandations: string;
}

interface StakeholderSatisfaction {
  id: number;
  name: string;
  c1: number; // Communication (1-5)
  c2: number; // Respect des délais (1-5)
  c3: number; // Qualité livrables (1-5)
  c4: number; // Accompagnement / Run (1-5)
}

interface DocumentArchivage {
  id: number;
  titre: string;
  archive: boolean;
  commentaire: string;
}

const DEFAULT_ECARTS: EcartItem[] = [
  {
    id: 'livrables',
    question: 'Les livrables ont-ils été livrés tels que prévus dans la charte de projet ?',
    reponse: 'Non',
    explication: "Quelques adaptations mineures de périmètre ont été actées en cours de projet sans impacter l'objectif global."
  },
  {
    id: 'objectifs',
    question: 'Les objectifs du projet, tels que définis dans la charte de projet, ont-ils été atteints ?',
    reponse: 'Oui',
    explication: "Tous les objectifs ont été atteints en terme de qualité, périmètre et budget. Il faudra attendre le décommissionnement complet des anciens systèmes avant de pouvoir bénéficier de tous les avantages des nouvelles solutions."
  },
  {
    id: 'delais',
    question: 'La date de fin du projet, telle qu’estimée dans la planification de base du plan de management de projet, a-t-elle été respectée ?',
    reponse: 'Oui',
    explication: "Le projet avait une date initiale du 29 novembre 2022 pour le basculement du système. Un plan B avait été développé pour permettre un basculement en décembre en cas de problème majeur. Ce plan de secours n'a pas été nécessaire et nous avons livré à la date initialement prévue."
  },
  {
    id: 'cout',
    question: 'Le coût du projet, tel qu’estimé dans la planification de base du plan de management de projet, a-t-il été respecté ?',
    reponse: 'Oui',
    explication: "Le périmètre et la date de démarrage ayant été respectés, le coût du projet a lui aussi été respecté. Il faut noter que certains postes de travail ont dû être mis à jour, et cela a été effectué par le service informatique et imputé au budget infrastructure de celui-ci."
  },
  {
    id: 'efforts',
    question: 'Les efforts du projet (en j-h), tels qu’estimés dans la planification de base ont-ils été respectés ?',
    reponse: 'Non',
    explication: "La partie Master Data du projet a été plus difficile que prévue et les corrections extrêmement importantes. La charge de travail de l'équipe Master Data a été dépassée de 5%. Ces efforts supplémentaires n'ont toutefois pas eu d'effet négatif sur les budgets."
  }
];

const DEFAULT_RISK_REVIEWS: RiskReviewItem[] = [
  {
    id: 1,
    risque: 'Qualité des données médiocre',
    impactPlanifie: 'Assignation ressources Master Data',
    impactReel: 'Corrections beaucoup plus importantes que prévues'
  },
  {
    id: 2,
    risque: 'Importance des immobilisations',
    impactPlanifie: "Peu d'impact",
    impactReel: 'Migration importante. Règles multiples'
  },
  {
    id: 3,
    risque: 'Défis de communication',
    impactPlanifie: 'Nécessité d’un responsable du changement',
    impactReel: 'Très nombreuses parties prenantes, parfois non identifiées'
  },
  {
    id: 4,
    risque: 'Utilisation serveurs du siège',
    impactPlanifie: 'Serveurs suffisants',
    impactReel: 'Nécessité de mettre en place le plan de contingence'
  },
  {
    id: 5,
    risque: 'Postes de travail à remplacer',
    impactPlanifie: 'Parc à jour',
    impactReel: 'Quelques postes avaient été oubliés. Problème corrigé.'
  }
];

const DEFAULT_REX: RexItem[] = [
  {
    domaine: 'Communication',
    lecons: 'Importance de tenir compte de la perspective des parties prenantes. L’intranet n’est pas accessible par/pour tout le monde.',
    recommandations: 'Sonder l’ensemble du personnel et comprendre les perceptions. Utiliser TOUS les panneaux d’information, y compris dans l’usine et à l’extérieur.'
  },
  {
    domaine: 'Contenu et intégration',
    lecons: 'Une usine d’assemblage est administrativement et fonctionnellement plus complexe.',
    recommandations: 'Garder toute la documentation de ce projet pour la prochaine mise à niveau SAP prévue l’année prochaine.'
  },
  {
    domaine: 'Achats / Contrats',
    lecons: 'Grosses différences de traitement entre petits/moyens et très gros clients.',
    recommandations: 'Les nouveaux paramétrages devraient permettre de faire face à ce défi avec agilité.'
  },
  {
    domaine: 'Coûts',
    lecons: 'Importance de prévoir l’imprévisible sur les interfaces techniques.',
    recommandations: 'Prévoir systématiquement un coussin de sécurité budgétaire de 5 à 10% pour l’infra.'
  },
  {
    domaine: 'Délais',
    lecons: 'Les périodes de congés d’été sont critiques pour les validations d’équipes.',
    recommandations: 'Prévoir des ressources de secours ou décaler les phases de recette hors vacances scolaires.'
  },
  {
    domaine: 'Qualité',
    lecons: 'Qualité parfois médiocre des données sources d’anciens ERPs.',
    recommandations: 'Faire une pré-étude de qualité des données dès la phase de cadrage sans sous-estimer l’impact.'
  },
  {
    domaine: 'Ressources humaines',
    lecons: 'L’absentéisme n’est pas une science exacte et pèse sur les profils clés.',
    recommandations: 'Prévoir le remplacement rapide et des doublons sur les postes d’utilisateurs référents.'
  },
  {
    domaine: 'Risques',
    lecons: 'La revue régulière des risques en comité hebdomadaire est primordiale.',
    recommandations: 'Ne jamais retirer cette revue de l’agenda des réunions de pilotage.'
  },
  {
    domaine: 'Technologies',
    lecons: 'Inventaires imparfaits sur les postes d’atelier distants.',
    recommandations: 'Ne pas oublier les coins les plus éloignés du centre de montage. Mettre à jour l’inventaire physique.'
  }
];

const DEFAULT_STAKEHOLDERS: StakeholderSatisfaction[] = [
  { id: 1, name: 'Partie prenante 1 (Direction Générale / Sponsor)', c1: 4, c2: 2, c3: 4, c4: 5 },
  { id: 2, name: 'Partie prenante 2 (Métier / Finance & Ventes)', c1: 3, c2: 3, c3: 5, c4: 3 },
  { id: 3, name: 'Partie prenante 3 (DSI / Équipe Technique)', c1: 2, c2: 4, c3: 4, c4: 4 },
  { id: 4, name: 'Partie prenante 4 (Utilisateurs clés & Exploitation)', c1: 3, c2: 3, c3: 3, c4: 5 },
  { id: 5, name: 'Partie prenante 5 (Partenaires & Intégrateur SAP)', c1: 3, c2: 4, c3: 4, c4: 3 }
];

const DEFAULT_DOCUMENTS: DocumentArchivage[] = [
  { id: 1, titre: "Étude d'opportunité", archive: true, commentaire: 'Archivé au format PDF certifié' },
  { id: 2, titre: 'Charte de projet (Project Charter)', archive: true, commentaire: 'Signé par le Sponsor' },
  { id: 3, titre: 'Cahier des charges fonctionnel', archive: true, commentaire: 'Version finale validée' },
  { id: 4, titre: 'Plan de management de projet (PMP)', archive: true, commentaire: 'Versions V1.0 à V2.3 archivées' },
  { id: 5, titre: 'Rapports d’avancement & Flash hebdomadaires', archive: true, commentaire: 'Ensemble des 24 rapports sauvegardés' },
  { id: 6, titre: 'Minutes de réunions & Comités de Pilotage', archive: true, commentaire: 'Comptes-rendus validés' },
  { id: 7, titre: 'Registre des problèmes et des risques', archive: true, commentaire: 'Clôturé avec matrice d’impact' },
  { id: 8, titre: 'Registre des livrables et PV de recette', archive: true, commentaire: 'PVs signés par la recette métier' },
  { id: 9, titre: 'Cahier de recette & Plan de tests', archive: true, commentaire: '100% scénarios validés' },
  { id: 10, titre: 'Bilan de projet & REX', archive: true, commentaire: 'À faire approuver par le comité final' },
  { id: 11, titre: 'Rapport financier définitif & Suivi VA', archive: true, commentaire: 'Clôture comptable arrêtée' },
  { id: 12, titre: 'Stratégie de reprise et passage en RUN', archive: true, commentaire: 'Transfert effectué vers équipe support' }
];

export const BilanProjetMatrix: React.FC<BilanProjetMatrixProps> = ({
  projectId,
  fileId,
  currentDossierId,
  activeSubTab = 'bilan'
}) => {
  const dossierSubKey = currentDossierId ? `-${currentDossierId}` : '';
  const storageKey = `faab-bilan-projet-matrix-${projectId}-${fileId}${dossierSubKey}`;

  const [activeTab, setActiveTab] = useState<string>(activeSubTab || 'bilan');

  // Identification
  const [ident, setIdent] = useState({
    nomCode: 'SAP30 / CMRL',
    reference: 'Stratégie 2022/CMRL',
    chefProjet: 'Chantal Dupont',
    organisation: 'Confassis S.A.',
    dateCreation: '15/03/2022',
    dateCloture: '29/11/2022',
    version: 'V2.3',
    docsJoints: 'Cahier de recette et minutes du Comité de Direction Final.'
  });

  // Ecarts
  const [ecarts, setEcarts] = useState<EcartItem[]>(() => {
    const saved = localStorage.getItem(`${storageKey}-ecarts`);
    return saved ? JSON.parse(saved) : DEFAULT_ECARTS;
  });

  // Risk reviews
  const [riskReviews, setRiskReviews] = useState<RiskReviewItem[]>(() => {
    const saved = localStorage.getItem(`${storageKey}-risks`);
    return saved ? JSON.parse(saved) : DEFAULT_RISK_REVIEWS;
  });

  // Unforeseen impacts & Post implementation text
  const [impactsImprevus, setImpactsImprevus] = useState<string>(() => {
    const saved = localStorage.getItem(`${storageKey}-unforeseen`);
    return saved || "Un absentéisme plus important que la normale nous a obligés à assigner des ressources de réserve pendant la période de paramétrage. Heureusement, les équipes du centre de compétence du siège ont pu déléguer des spécialistes très rapidement, réduisant ainsi l'impact du problème.";
  });

  const [postImplementation, setPostImplementation] = useState<string>(() => {
    const saved = localStorage.getItem(`${storageKey}-postimpl`);
    return saved || "Certaines erreurs ont été détectées pendant la clôture de novembre. Ces problèmes ont été immédiatement corrigés et la clôture de décembre s'est passée sans encombres. Le passage en RUN est stabilisé.";
  });

  // REX
  const [rexItems, setRexItems] = useState<RexItem[]>(() => {
    const saved = localStorage.getItem(`${storageKey}-rex`);
    return saved ? JSON.parse(saved) : DEFAULT_REX;
  });

  // Satisfaction Matrix
  const [stakeholders, setStakeholders] = useState<StakeholderSatisfaction[]>(() => {
    const saved = localStorage.getItem(`${storageKey}-satisfaction`);
    return saved ? JSON.parse(saved) : DEFAULT_STAKEHOLDERS;
  });

  // Documents
  const [documents, setDocuments] = useState<DocumentArchivage[]>(() => {
    const saved = localStorage.getItem(`${storageKey}-docs`);
    return saved ? JSON.parse(saved) : DEFAULT_DOCUMENTS;
  });

  // Cloture validation
  const [clotureSignee, setClotureSignee] = useState<boolean>(() => {
    const saved = localStorage.getItem(`${storageKey}-cloture`);
    return saved === 'true';
  });

  useEffect(() => {
    if (activeSubTab) setActiveTab(activeSubTab);
  }, [activeSubTab]);

  // Persistence
  useEffect(() => {
    localStorage.setItem(`${storageKey}-ecarts`, JSON.stringify(ecarts));
  }, [ecarts, storageKey]);

  useEffect(() => {
    localStorage.setItem(`${storageKey}-risks`, JSON.stringify(riskReviews));
  }, [riskReviews, storageKey]);

  useEffect(() => {
    localStorage.setItem(`${storageKey}-unforeseen`, impactsImprevus);
  }, [impactsImprevus, storageKey]);

  useEffect(() => {
    localStorage.setItem(`${storageKey}-postimpl`, postImplementation);
  }, [postImplementation, storageKey]);

  useEffect(() => {
    localStorage.setItem(`${storageKey}-rex`, JSON.stringify(rexItems));
  }, [rexItems, storageKey]);

  useEffect(() => {
    localStorage.setItem(`${storageKey}-satisfaction`, JSON.stringify(stakeholders));
  }, [stakeholders, storageKey]);

  useEffect(() => {
    localStorage.setItem(`${storageKey}-docs`, JSON.stringify(documents));
  }, [documents, storageKey]);

  useEffect(() => {
    localStorage.setItem(`${storageKey}-cloture`, String(clotureSignee));
  }, [clotureSignee, storageKey]);

  // Calculation of satisfaction
  const satisfactionAverages = useMemo(() => {
    const rowAvgs = stakeholders.map(s => {
      const avg = (s.c1 + s.c2 + s.c3 + s.c4) / 4;
      return parseFloat(avg.toFixed(2));
    });
    const globalAvg = rowAvgs.length > 0 ? rowAvgs.reduce((a, b) => a + b, 0) / rowAvgs.length : 0;
    return {
      rowAvgs,
      globalAvg: parseFloat(globalAvg.toFixed(2))
    };
  }, [stakeholders]);

  return (
    <div style={{ flex: 1, display: 'flex', flexDirection: 'column', background: 'var(--main-bg)', color: 'var(--main-text)', overflowY: 'auto' }}>
      {/* Top Header */}
      <div style={{ padding: '1.25rem 2rem', background: 'var(--main-surface)', borderBottom: '1px solid var(--main-border)', display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
        <div>
          <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
            <span style={{ fontSize: '1.3rem' }}>🏆</span>
            <h1 style={{ fontSize: '1.25rem', fontWeight: 900, margin: 0, color: 'var(--main-text)' }}>
              34 — Bilan de Projet & REX
            </h1>
            <span style={{ padding: '4px 10px', background: 'rgba(16, 185, 129, 0.12)', color: '#10b981', borderRadius: '20px', fontSize: '0.72rem', fontWeight: 800 }}>
              EXCEL 34 DYNAMIQUE
            </span>
          </div>
          <p style={{ margin: '4px 0 0', fontSize: '0.8rem', color: 'var(--text-muted)' }}>
            Kit Chef de Projet • Analyse des écarts, Retour d'expérience (REX 9 domaines), Matrice de satisfaction et Clôture RUN
          </p>
        </div>
      </div>

      {/* Sub-tabs bar */}
      <div style={{ display: 'flex', gap: '8px', padding: '0.75rem 2rem', background: 'var(--main-surface)', borderBottom: '1px solid var(--main-border)', overflowX: 'auto' }}>
        {[
          { id: 'bilan', label: '📋 Bilan & Écarts (Sec. 1 à 5)' },
          { id: 'rex', label: '💡 Leçons Apprises (REX 9 Domaines)' },
          { id: 'satisfaction', label: '⭐ Matrice de Satisfaction (Sec. 7)' },
          { id: 'cloture', label: '📦 Dépôt des Documents & Clôture (Sec. 8 & 9)' }
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

      <div style={{ padding: '2rem', display: 'flex', flexDirection: 'column', gap: '2rem' }}>
        {/* TAB 1: BILAN ET ÉCARTS */}
        {activeTab === 'bilan' && (
          <>
            {/* 1. IDENTIFICATION */}
            <div style={{ background: 'var(--main-surface)', padding: '2rem', borderRadius: '20px', border: '1px solid var(--main-border)', boxShadow: 'var(--card-shadow)' }}>
              <h3 style={{ margin: '0 0 1.25rem', fontSize: '1rem', fontWeight: 900, color: 'var(--excel-color)', textTransform: 'uppercase' }}>
                1. Identification du Projet
              </h3>
              <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(220px, 1fr))', gap: '1.25rem' }}>
                <div>
                  <label style={{ display: 'block', fontSize: '0.75rem', fontWeight: 800, color: 'var(--text-muted)', marginBottom: '4px' }}>Nom / Code projet</label>
                  <input
                    type="text"
                    value={ident.nomCode}
                    onChange={(e) => setIdent({ ...ident, nomCode: e.target.value })}
                    style={{ width: '100%', padding: '8px 12px', borderRadius: '8px', border: '1px solid var(--main-border)', background: 'var(--main-bg)', color: 'var(--main-text)', fontWeight: 700 }}
                  />
                </div>
                <div>
                  <label style={{ display: 'block', fontSize: '0.75rem', fontWeight: 800, color: 'var(--text-muted)', marginBottom: '4px' }}>Référence</label>
                  <input
                    type="text"
                    value={ident.reference}
                    onChange={(e) => setIdent({ ...ident, reference: e.target.value })}
                    style={{ width: '100%', padding: '8px 12px', borderRadius: '8px', border: '1px solid var(--main-border)', background: 'var(--main-bg)', color: 'var(--main-text)' }}
                  />
                </div>
                <div>
                  <label style={{ display: 'block', fontSize: '0.75rem', fontWeight: 800, color: 'var(--text-muted)', marginBottom: '4px' }}>Chef de projet</label>
                  <input
                    type="text"
                    value={ident.chefProjet}
                    onChange={(e) => setIdent({ ...ident, chefProjet: e.target.value })}
                    style={{ width: '100%', padding: '8px 12px', borderRadius: '8px', border: '1px solid var(--main-border)', background: 'var(--main-bg)', color: 'var(--main-text)', fontWeight: 700 }}
                  />
                </div>
                <div>
                  <label style={{ display: 'block', fontSize: '0.75rem', fontWeight: 800, color: 'var(--text-muted)', marginBottom: '4px' }}>Organisation / Service</label>
                  <input
                    type="text"
                    value={ident.organisation}
                    onChange={(e) => setIdent({ ...ident, organisation: e.target.value })}
                    style={{ width: '100%', padding: '8px 12px', borderRadius: '8px', border: '1px solid var(--main-border)', background: 'var(--main-bg)', color: 'var(--main-text)' }}
                  />
                </div>
                <div>
                  <label style={{ display: 'block', fontSize: '0.75rem', fontWeight: 800, color: 'var(--text-muted)', marginBottom: '4px' }}>Date Clôture / Livraison</label>
                  <input
                    type="text"
                    value={ident.dateCloture}
                    onChange={(e) => setIdent({ ...ident, dateCloture: e.target.value })}
                    style={{ width: '100%', padding: '8px 12px', borderRadius: '8px', border: '1px solid var(--main-border)', background: 'var(--main-bg)', color: 'var(--main-text)' }}
                  />
                </div>
                <div>
                  <label style={{ display: 'block', fontSize: '0.75rem', fontWeight: 800, color: 'var(--text-muted)', marginBottom: '4px' }}>Version</label>
                  <input
                    type="text"
                    value={ident.version}
                    onChange={(e) => setIdent({ ...ident, version: e.target.value })}
                    style={{ width: '100%', padding: '8px 12px', borderRadius: '8px', border: '1px solid var(--main-border)', background: 'var(--main-bg)', color: 'var(--main-text)', fontWeight: 800 }}
                  />
                </div>
              </div>
            </div>

            {/* 2. ANALYSE DES ÉCARTS (5 questions) */}
            <div style={{ background: 'var(--main-surface)', padding: '2rem', borderRadius: '20px', border: '1px solid var(--main-border)', boxShadow: 'var(--card-shadow)' }}>
              <h3 style={{ margin: '0 0 1.5rem', fontSize: '1rem', fontWeight: 900, color: '#10b981', textTransform: 'uppercase' }}>
                2. Analyse Critique des Écarts
              </h3>
              <div style={{ display: 'flex', flexDirection: 'column', gap: '1.25rem' }}>
                {ecarts.map((item, idx) => (
                  <div
                    key={item.id}
                    style={{
                      padding: '1.25rem',
                      borderRadius: '14px',
                      background: 'var(--main-bg)',
                      border: '1px solid var(--main-border)',
                      display: 'flex',
                      flexDirection: 'column',
                      gap: '10px'
                    }}
                  >
                    <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                      <span style={{ fontWeight: 800, fontSize: '0.92rem', color: 'var(--main-text)' }}>
                        {idx + 1}. {item.question}
                      </span>
                      <div style={{ display: 'flex', gap: '6px' }}>
                        <button
                          onClick={() => {
                            const next = [...ecarts];
                            next[idx].reponse = 'Oui';
                            setEcarts(next);
                          }}
                          style={{
                            padding: '4px 14px',
                            borderRadius: '20px',
                            border: item.reponse === 'Oui' ? '1px solid #10b981' : '1px solid var(--main-border)',
                            background: item.reponse === 'Oui' ? '#10b981' : 'var(--main-surface)',
                            color: item.reponse === 'Oui' ? '#fff' : 'var(--text-muted)',
                            fontWeight: 900,
                            fontSize: '0.75rem',
                            cursor: 'pointer'
                          }}
                        >
                          Oui
                        </button>
                        <button
                          onClick={() => {
                            const next = [...ecarts];
                            next[idx].reponse = 'Non';
                            setEcarts(next);
                          }}
                          style={{
                            padding: '4px 14px',
                            borderRadius: '20px',
                            border: item.reponse === 'Non' ? '1px solid #ef4444' : '1px solid var(--main-border)',
                            background: item.reponse === 'Non' ? '#ef4444' : 'var(--main-surface)',
                            color: item.reponse === 'Non' ? '#fff' : 'var(--text-muted)',
                            fontWeight: 900,
                            fontSize: '0.75rem',
                            cursor: 'pointer'
                          }}
                        >
                          Non
                        </button>
                      </div>
                    </div>
                    <textarea
                      rows={2}
                      value={item.explication}
                      onChange={(e) => {
                        const next = [...ecarts];
                        next[idx].explication = e.target.value;
                        setEcarts(next);
                      }}
                      placeholder="Commentaires et explications détaillées..."
                      style={{
                        width: '100%',
                        padding: '8px 12px',
                        borderRadius: '8px',
                        border: '1px solid var(--main-border)',
                        background: 'var(--main-surface)',
                        color: 'var(--main-text)',
                        fontSize: '0.85rem',
                        lineHeight: '1.5',
                        resize: 'vertical'
                      }}
                    />
                  </div>
                ))}
              </div>
            </div>

            {/* 3. RETOUR SUR L'ANALYSE DES RISQUES */}
            <div style={{ background: 'var(--main-surface)', borderRadius: '20px', border: '1px solid var(--main-border)', overflow: 'hidden', boxShadow: 'var(--card-shadow)' }}>
              <div style={{ background: 'var(--excel-color)', padding: '16px 24px', display: 'flex', justifyContent: 'space-between', alignItems: 'center', color: '#fff' }}>
                <span style={{ fontWeight: 900, fontSize: '0.9rem', textTransform: 'uppercase' }}>
                  3. Retour sur l'Analyse des Risques Survenus
                </span>
                <button
                  onClick={() => {
                    const nextId = riskReviews.length + 1;
                    setRiskReviews([...riskReviews, { id: nextId, risque: 'Nouveau risque', impactPlanifie: 'Planifié', impactReel: 'Réel' }]);
                  }}
                  style={{
                    padding: '6px 14px',
                    background: 'rgba(255, 255, 255, 0.2)',
                    color: '#fff',
                    border: '1px solid rgba(255, 255, 255, 0.4)',
                    borderRadius: '8px',
                    fontWeight: 800,
                    fontSize: '0.75rem',
                    cursor: 'pointer'
                  }}
                >
                  + Ajouter une ligne
                </button>
              </div>

              <div style={{ overflowX: 'auto' }}>
                <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '0.88rem' }}>
                  <thead>
                    <tr style={{ background: 'rgba(0,0,0,0.02)', borderBottom: '2px solid var(--main-border)', color: 'var(--text-muted)' }}>
                      <th style={{ padding: '12px 16px', textAlign: 'left', width: '32%' }}>Risque Survenu</th>
                      <th style={{ padding: '12px 16px', textAlign: 'left', width: '34%' }}>Impact Planifié</th>
                      <th style={{ padding: '12px 16px', textAlign: 'left', width: '30%' }}>Impact Réel</th>
                      <th style={{ padding: '12px 16px', textAlign: 'center', width: '4%' }}>Action</th>
                    </tr>
                  </thead>
                  <tbody>
                    {riskReviews.map((item, idx) => (
                      <tr key={item.id} style={{ borderBottom: '1px solid var(--main-border)' }}>
                        <td style={{ padding: '10px 16px' }}>
                          <input
                            type="text"
                            value={item.risque}
                            onChange={(e) => {
                              const next = [...riskReviews];
                              next[idx].risque = e.target.value;
                              setRiskReviews(next);
                            }}
                            style={{ width: '100%', background: 'transparent', border: '1px solid transparent', color: 'var(--main-text)', fontWeight: 700, padding: '4px 6px', borderRadius: '4px' }}
                          />
                        </td>
                        <td style={{ padding: '10px 16px' }}>
                          <input
                            type="text"
                            value={item.impactPlanifie}
                            onChange={(e) => {
                              const next = [...riskReviews];
                              next[idx].impactPlanifie = e.target.value;
                              setRiskReviews(next);
                            }}
                            style={{ width: '100%', background: 'transparent', border: '1px solid transparent', color: 'var(--text-muted)', padding: '4px 6px', borderRadius: '4px' }}
                          />
                        </td>
                        <td style={{ padding: '10px 16px' }}>
                          <input
                            type="text"
                            value={item.impactReel}
                            onChange={(e) => {
                              const next = [...riskReviews];
                              next[idx].impactReel = e.target.value;
                              setRiskReviews(next);
                            }}
                            style={{ width: '100%', background: 'transparent', border: '1px solid transparent', color: '#ef4444', fontWeight: 600, padding: '4px 6px', borderRadius: '4px' }}
                          />
                        </td>
                        <td style={{ padding: '10px 16px', textAlign: 'center' }}>
                          <button
                            onClick={() => setRiskReviews(riskReviews.filter((_, i) => i !== idx))}
                            style={{ background: 'transparent', border: 'none', color: '#ef4444', cursor: 'pointer' }}
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

            {/* 4 & 5: IMPACTS IMPRÉVUS & POST IMPLÉMENTATION */}
            <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '2rem' }}>
              <div style={{ background: 'var(--main-surface)', borderRadius: '20px', border: '1px solid var(--main-border)', padding: '2rem', boxShadow: 'var(--card-shadow)' }}>
                <h3 style={{ margin: '0 0 1rem', fontSize: '0.95rem', fontWeight: 900, color: '#f59e0b', textTransform: 'uppercase' }}>
                  4. Impacts Imprévus non Planifiés
                </h3>
                <textarea
                  rows={4}
                  value={impactsImprevus}
                  onChange={(e) => setImpactsImprevus(e.target.value)}
                  style={{ width: '100%', padding: '10px', borderRadius: '10px', border: '1px solid var(--main-border)', background: 'var(--main-bg)', color: 'var(--main-text)', fontSize: '0.85rem', lineHeight: '1.5' }}
                />
              </div>

              <div style={{ background: 'var(--main-surface)', borderRadius: '20px', border: '1px solid var(--main-border)', padding: '2rem', boxShadow: 'var(--card-shadow)' }}>
                <h3 style={{ margin: '0 0 1rem', fontSize: '0.95rem', fontWeight: 900, color: '#3b82f6', textTransform: 'uppercase' }}>
                  5. Étape de Post-Implémentation (Stabilisation)
                </h3>
                <textarea
                  rows={4}
                  value={postImplementation}
                  onChange={(e) => setPostImplementation(e.target.value)}
                  style={{ width: '100%', padding: '10px', borderRadius: '10px', border: '1px solid var(--main-border)', background: 'var(--main-bg)', color: 'var(--main-text)', fontSize: '0.85rem', lineHeight: '1.5' }}
                />
              </div>
            </div>
          </>
        )}

        {/* TAB 2: REX (LEÇONS APPRISES SUR 9 DOMAINES) */}
        {activeTab === 'rex' && (
          <div style={{ background: 'var(--main-surface)', borderRadius: '20px', border: '1px solid var(--main-border)', overflow: 'hidden', boxShadow: 'var(--card-shadow)' }}>
            <div style={{ background: 'var(--excel-color)', padding: '16px 24px', display: 'flex', justifyContent: 'space-between', alignItems: 'center', color: '#fff' }}>
              <div>
                <span style={{ fontWeight: 900, fontSize: '0.95rem', textTransform: 'uppercase' }}>
                  6. Leçons Apprises et Recommandations (REX 9 Domaines)
                </span>
                <p style={{ margin: '4px 0 0', fontSize: '0.75rem', opacity: 0.9 }}>
                  Capitalisation transversale sur les processus, outils, relations et pilotage
                </p>
              </div>
            </div>

            <div style={{ overflowX: 'auto' }}>
              <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '0.88rem' }}>
                <thead>
                  <tr style={{ background: 'rgba(0,0,0,0.02)', borderBottom: '2px solid var(--main-border)', color: 'var(--text-muted)' }}>
                    <th style={{ padding: '12px 18px', textAlign: 'left', width: '20%' }}>Domaine</th>
                    <th style={{ padding: '12px 18px', textAlign: 'left', width: '40%' }}>Leçons Apprises (Ce qui a fonctionné / dysfonctionné)</th>
                    <th style={{ padding: '12px 18px', textAlign: 'left', width: '40%' }}>Recommandations pour Projets Futurs</th>
                  </tr>
                </thead>
                <tbody>
                  {rexItems.map((item, idx) => (
                    <tr key={idx} style={{ borderBottom: '1px solid var(--main-border)' }}>
                      <td style={{ padding: '14px 18px', fontWeight: 900, color: '#10b981', verticalAlign: 'top' }}>
                        {item.domaine}
                      </td>
                      <td style={{ padding: '12px 18px', verticalAlign: 'top' }}>
                        <textarea
                          rows={2}
                          value={item.lecons}
                          onChange={(e) => {
                            const next = [...rexItems];
                            next[idx].lecons = e.target.value;
                            setRexItems(next);
                          }}
                          style={{ width: '100%', background: 'transparent', border: '1px solid transparent', color: 'var(--main-text)', fontSize: '0.85rem', lineHeight: '1.5', resize: 'vertical' }}
                        />
                      </td>
                      <td style={{ padding: '12px 18px', verticalAlign: 'top' }}>
                        <textarea
                          rows={2}
                          value={item.recommandations}
                          onChange={(e) => {
                            const next = [...rexItems];
                            next[idx].recommandations = e.target.value;
                            setRexItems(next);
                          }}
                          style={{ width: '100%', background: 'transparent', border: '1px solid transparent', color: 'var(--text-muted)', fontSize: '0.85rem', lineHeight: '1.5', resize: 'vertical' }}
                        />
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>
        )}

        {/* TAB 3: MATRICE DE SATISFACTION */}
        {activeTab === 'satisfaction' && (
          <div style={{ display: 'flex', flexDirection: 'column', gap: '2rem' }}>
            {/* SCORE HERO BANNER */}
            <div style={{
              background: 'linear-gradient(135deg, rgba(16, 185, 129, 0.12), rgba(59, 130, 246, 0.12))',
              borderRadius: '24px',
              border: '1px solid var(--main-border)',
              padding: '2.5rem',
              display: 'flex',
              justifyContent: 'space-between',
              alignItems: 'center',
              boxShadow: 'var(--card-shadow)'
            }}>
              <div>
                <span style={{ fontSize: '0.8rem', fontWeight: 900, color: '#10b981', textTransform: 'uppercase', letterSpacing: '0.1em' }}>
                  Matrice d'Évaluation de Fin de Projet
                </span>
                <h2 style={{ fontSize: '1.8rem', fontWeight: 950, margin: '8px 0 4px', color: 'var(--main-text)' }}>
                  Satisfaction Générale Consolidée
                </h2>
                <p style={{ margin: 0, fontSize: '0.9rem', color: 'var(--text-muted)' }}>
                  Moyenne calculée automatiquement sur l'ensemble des 5 parties prenantes et des 4 critères d'évaluation
                </p>
              </div>

              <div style={{ textAlign: 'center', background: 'var(--main-surface)', padding: '1.5rem 2.5rem', borderRadius: '20px', border: '1px solid var(--main-border)', boxShadow: 'var(--card-shadow)' }}>
                <div style={{ fontSize: '3.5rem', fontWeight: 950, color: '#10b981', lineHeight: 1 }}>
                  {satisfactionAverages.globalAvg} <span style={{ fontSize: '1.4rem', opacity: 0.5, color: 'var(--text-muted)' }}>/ 5</span>
                </div>
                <div style={{ marginTop: '8px', fontSize: '0.85rem', fontWeight: 800, color: '#f59e0b' }}>
                  {'★'.repeat(Math.round(satisfactionAverages.globalAvg))}
                  {'☆'.repeat(5 - Math.round(satisfactionAverages.globalAvg))}
                </div>
                <div style={{ marginTop: '4px', fontSize: '0.72rem', fontWeight: 800, color: 'var(--text-muted)', textTransform: 'uppercase' }}>
                  Score Global Excel
                </div>
              </div>
            </div>

            {/* SATISFACTION MATRIX TABLE */}
            <div style={{ background: 'var(--main-surface)', borderRadius: '20px', border: '1px solid var(--main-border)', overflow: 'hidden', boxShadow: 'var(--card-shadow)' }}>
              <div style={{ background: 'var(--excel-color)', padding: '16px 24px', color: '#fff', fontWeight: 900, fontSize: '0.9rem', textTransform: 'uppercase' }}>
                7. Évaluation Détaillée par Partie Prenante (Barème de 1 à 5)
              </div>
              <div style={{ overflowX: 'auto' }}>
                <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '0.88rem' }}>
                  <thead>
                    <tr style={{ background: 'rgba(0,0,0,0.02)', borderBottom: '2px solid var(--main-border)', color: 'var(--text-muted)' }}>
                      <th style={{ padding: '14px 18px', textAlign: 'left' }}>Partie Prenante</th>
                      <th style={{ padding: '14px 18px', textAlign: 'center', width: '130px' }}>Critère 1<br /><span style={{ fontSize: '0.72rem', fontWeight: 600 }}>(Communication)</span></th>
                      <th style={{ padding: '14px 18px', textAlign: 'center', width: '130px' }}>Critère 2<br /><span style={{ fontSize: '0.72rem', fontWeight: 600 }}>(Respect Délais)</span></th>
                      <th style={{ padding: '14px 18px', textAlign: 'center', width: '130px' }}>Critère 3<br /><span style={{ fontSize: '0.72rem', fontWeight: 600 }}>(Qualité Livrables)</span></th>
                      <th style={{ padding: '14px 18px', textAlign: 'center', width: '130px' }}>Critère 4<br /><span style={{ fontSize: '0.72rem', fontWeight: 600 }}>(Accompagnement)</span></th>
                      <th style={{ padding: '14px 18px', textAlign: 'center', width: '140px', fontWeight: 900, color: 'var(--main-text)' }}>Satisfaction (Moyenne)</th>
                    </tr>
                  </thead>
                  <tbody>
                    {stakeholders.map((s, idx) => (
                      <tr key={s.id} style={{ borderBottom: '1px solid var(--main-border)' }}>
                        <td style={{ padding: '12px 18px', fontWeight: 800 }}>
                          <input
                            type="text"
                            value={s.name}
                            onChange={(e) => {
                              const next = [...stakeholders];
                              next[idx].name = e.target.value;
                              setStakeholders(next);
                            }}
                            style={{ width: '100%', background: 'transparent', border: '1px solid transparent', color: 'var(--main-text)', fontWeight: 800, padding: '4px 6px', borderRadius: '4px' }}
                          />
                        </td>
                        {(['c1', 'c2', 'c3', 'c4'] as const).map(cKey => (
                          <td key={cKey} style={{ padding: '12px 18px', textAlign: 'center' }}>
                            <select
                              value={s[cKey]}
                              onChange={(e) => {
                                const next = [...stakeholders];
                                next[idx][cKey] = parseInt(e.target.value) || 1;
                                setStakeholders(next);
                              }}
                              style={{
                                padding: '6px 12px',
                                borderRadius: '8px',
                                border: '1px solid var(--main-border)',
                                background: 'var(--main-bg)',
                                color: 'var(--main-text)',
                                fontWeight: 800,
                                cursor: 'pointer'
                              }}
                            >
                              {[1, 2, 3, 4, 5].map(val => (
                                <option key={val} value={val}>{val} ★</option>
                              ))}
                            </select>
                          </td>
                        ))}
                        <td style={{ padding: '12px 18px', textAlign: 'center', fontWeight: 950, fontSize: '1rem', color: '#10b981' }}>
                          {satisfactionAverages.rowAvgs[idx]} / 5
                        </td>
                      </tr>
                    ))}
                    <tr style={{ background: 'rgba(16, 185, 129, 0.06)', fontWeight: 950 }}>
                      <td style={{ padding: '16px 18px', textTransform: 'uppercase', color: 'var(--main-text)' }}>
                        Satisfaction Générale Consolidée :
                      </td>
                      <td colSpan={4} style={{ textAlign: 'right', padding: '16px 18px', color: 'var(--text-muted)' }}>
                        Formule = AVERAGE(lignes 1 à 5) :
                      </td>
                      <td style={{ padding: '16px 18px', textAlign: 'center', fontSize: '1.2rem', color: '#10b981' }}>
                        {satisfactionAverages.globalAvg} / 5
                      </td>
                    </tr>
                  </tbody>
                </table>
              </div>
            </div>
          </div>
        )}

        {/* TAB 4: DÉPÔT & CLÔTURE */}
        {activeTab === 'cloture' && (
          <div style={{ display: 'flex', flexDirection: 'column', gap: '2rem' }}>
            {/* 8. DEPOT DES DOCUMENTS */}
            <div style={{ background: 'var(--main-surface)', borderRadius: '20px', border: '1px solid var(--main-border)', overflow: 'hidden', boxShadow: 'var(--card-shadow)' }}>
              <div style={{ background: 'var(--excel-color)', padding: '16px 24px', color: '#fff', fontWeight: 900, fontSize: '0.9rem', textTransform: 'uppercase' }}>
                8. Dépôt et Archivage des Documents du Projet (12 Documents Clés)
              </div>
              <div style={{ overflowX: 'auto' }}>
                <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '0.88rem' }}>
                  <thead>
                    <tr style={{ background: 'rgba(0,0,0,0.02)', borderBottom: '2px solid var(--main-border)', color: 'var(--text-muted)' }}>
                      <th style={{ padding: '12px 18px', textAlign: 'left', width: '40%' }}>Type de Document</th>
                      <th style={{ padding: '12px 18px', textAlign: 'center', width: '120px' }}>Archivé ?</th>
                      <th style={{ padding: '12px 18px', textAlign: 'left' }}>Commentaires & Emplacement</th>
                    </tr>
                  </thead>
                  <tbody>
                    {documents.map((doc, idx) => (
                      <tr key={doc.id} style={{ borderBottom: '1px solid var(--main-border)' }}>
                        <td style={{ padding: '12px 18px', fontWeight: 800 }}>{doc.titre}</td>
                        <td style={{ padding: '12px 18px', textAlign: 'center' }}>
                          <input
                            type="checkbox"
                            checked={doc.archive}
                            onChange={(e) => {
                              const next = [...documents];
                              next[idx].archive = e.target.checked;
                              setDocuments(next);
                            }}
                            style={{ width: '18px', height: '18px', accentColor: '#10b981', cursor: 'pointer' }}
                          />
                        </td>
                        <td style={{ padding: '12px 18px' }}>
                          <input
                            type="text"
                            value={doc.commentaire}
                            onChange={(e) => {
                              const next = [...documents];
                              next[idx].commentaire = e.target.value;
                              setDocuments(next);
                            }}
                            style={{ width: '100%', background: 'transparent', border: '1px solid transparent', color: 'var(--text-muted)', padding: '4px 6px', borderRadius: '4px' }}
                          />
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>

            {/* 9. CLOTURE FORMELLE */}
            <div style={{
              background: 'var(--main-surface)',
              padding: '2.5rem',
              borderRadius: '24px',
              border: clotureSignee ? '2px solid #10b981' : '1px solid var(--main-border)',
              boxShadow: 'var(--card-shadow)',
              textAlign: 'center'
            }}>
              <div style={{ fontSize: '3rem', marginBottom: '1rem' }}>
                {clotureSignee ? '🎉' : '✍️'}
              </div>
              <h2 style={{ fontSize: '1.5rem', fontWeight: 950, color: 'var(--main-text)', margin: '0 0 1rem' }}>
                9. Clôture Formelle du Projet & Passage en RUN
              </h2>
              <p style={{ maxWidth: '750px', margin: '0 auto 2rem', fontSize: '0.95rem', lineHeight: '1.7', color: 'var(--text-muted)' }}>
                « Les parties prenantes attestent avoir reçu l'ensemble des livrables en conformité avec les attentes de la charte de projet.
                L'équipe support a été définie et formée. Le projet est réputé définitivement clos pour transfert aux opérations RUN. »
              </p>
              <button
                onClick={() => setClotureSignee(!clotureSignee)}
                style={{
                  padding: '14px 32px',
                  borderRadius: '30px',
                  border: 'none',
                  background: clotureSignee ? '#10b981' : 'linear-gradient(135deg, #3b82f6, #1d4ed8)',
                  color: '#fff',
                  fontWeight: 900,
                  fontSize: '0.95rem',
                  cursor: 'pointer',
                  boxShadow: '0 8px 16px rgba(0,0,0,0.15)',
                  transition: 'all 0.2s ease'
                }}
              >
                {clotureSignee ? '✅ Projet Officiellement Clôturé (Signé)' : 'Signer et Valider la Clôture Définitive'}
              </button>
            </div>
          </div>
        )}
      </div>
    </div>
  );
};
