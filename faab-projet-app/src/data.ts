export interface ProjectFile {
    name: string;
    type: 'excel' | 'word' | 'ppt' | 'other';
    id: string;
}

export interface TabData {
    id: string;
    title: string;
    files: ProjectFile[];
}

export const tabsData: TabData[] = [
    {
        id: 'avant-projet', title: 'Avant projet', files: [
            { id: 'selection-projets', name: 'Sélection de projets', type: 'excel' },
            { id: 'matrice-faisabilite', name: 'Matrice de faisabilité', type: 'excel' },
            { id: 'doc-etude-opportunite', name: "Étude d'Opportunité (Business Case)", type: 'word' }
        ]
    },
    {
        id: 'initialisation', title: 'Initialisation', files: [
            { id: 'doc-charte-projet', name: 'Charte de Projet', type: 'word' },
            { id: 'ppt-lancement', name: 'Support Formation : Réunion de Lancement', type: 'ppt' },
            { id: 'prenantes-1', name: 'Partis inscrits', type: 'excel' },
            { id: 'prenantes-2', name: 'Registre des décideurs', type: 'excel' },
            { id: 'macro-planning', name: 'Macro Planning', type: 'excel' },
            { id: 'risques-2', name: 'Registre des risques Variante 2', type: 'excel' },
            { id: 'risques-1', name: 'Registre risques et problèmes Variante 1', type: 'excel' }
        ]
    },
    {
        id: 'planification', title: 'Planification', files: [
            { id: 'doc-expression-besoins', name: 'Expression des Besoins', type: 'word' },
            { id: 'doc-cahier-des-charges', name: 'Cahier des Charges Fonctionnel', type: 'word' },
            { id: 'doc-poc', name: 'Dossier Proof of Concept (POC)', type: 'word' },
            { id: 'ppt-pert', name: 'Support Formation : Méthodologie Réseau PERT', type: 'ppt' },
            { id: 'matrice-exigences', name: 'Matrice des exigences', type: 'excel' },
            { id: 'wbs', name: 'WBS-PBS', type: 'excel' },
            { id: 'gantt', name: 'Planning projet Gantt', type: 'excel' },
            { id: 'competences', name: 'Matrice des compétences', type: 'excel' },
            { id: 'raci', name: 'Matrice RACI', type: 'excel' },
            { id: 'comm-fr', name: 'Actions de Communication FR', type: 'excel' },
            { id: 'comm-act-en', name: 'Communication Action Plan EN', type: 'excel' },
            { id: 'plan-charge', name: 'Plan de charge', type: 'excel' },
            { id: 'doc-plan-communication', name: 'Plan de Communication', type: 'word' },
            { id: 'doc-plan-projet', name: 'Plan de Management de Projet (PMP)', type: 'word' },
            { id: 'doc-plan-qualite', name: "Plan d'Assurance Qualité (PAQ)", type: 'word' },
            { id: 'doc-estimation-budget', name: "Rapport d'Estimation Budgétaire", type: 'word' }
        ]
    },
    {
        id: 'execution', title: 'Exécution', files: [
            { id: 'suivi-actions', name: 'Suivi des actions', type: 'excel' },
            { id: 'suivi-livrables', name: 'Suivi des livrables', type: 'excel' },
            { id: 'chg-req-reg', name: 'Registre demandes de changement', type: 'excel' },
            { id: 'ppt-workflow-changement', name: 'Support Formation : Workflow des Changements', type: 'ppt' },
            { id: 'doc-cr-reunion', name: 'Compte-Rendu Synthétique de Réunion', type: 'word' },
            { id: 'doc-minutes-copil', name: 'Procès-Verbal Réunion COPIL', type: 'word' },
            { id: 'amdec', name: 'AMDEC', type: 'excel' },
            { id: 'plan-test', name: 'Plan de test', type: 'excel' },
            { id: 'doc-strategie-tests', name: 'Stratégie de Recette et Tests', type: 'word' },
            { id: 'doc-cahier-recette', name: 'Cahier de Recette Utilisateur', type: 'word' }
        ]
    },
    {
        id: 'controle', title: 'Contrôle', files: [
            { id: 'suivi-budget', name: 'Suivi de budget', type: 'excel' },
            { id: 'tb', name: 'Tableau de bord', type: 'excel' },
            { id: 'rapport-flash-xls', name: 'Rapport Flash', type: 'excel' },
            { id: 'ppt-rapport-flash', name: 'Support Formation : Reporting Flash Projet', type: 'ppt' },
            { id: 'ppt-coproj', name: 'Support Formation : Conduite Comité Projet', type: 'ppt' },
            { id: 'ppt-copil', name: 'Support Formation : Conduite Comité Pilotage', type: 'ppt' }
        ]
    },
    {
        id: 'cloture', title: 'Clôture', files: [
            { id: 'bilan', name: 'Bilan de projet', type: 'excel' },
            { id: 'doc-pv-reception', name: 'PV de Réception des Livrables', type: 'word' }
        ]
    }
];
