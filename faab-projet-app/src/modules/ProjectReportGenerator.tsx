import React, { useState, useEffect, useRef } from 'react';

export interface ReportSection {
  id: string;
  title: string;
  content: string;
  type?: 'text' | 'table' | 'signatures';
  tableHeaders?: string[];
  tableRows?: string[][];
}

export interface ReportTemplate {
  id: string;
  name: string;
  category: string;
  code: string;
  title: string;
  subtitle: string;
  version: string;
  date: string;
  author: string;
  approver: string;
  status: 'Brouillon' | 'En revue' | 'Validé' | 'Diffusé';
  summary: string;
  sections: ReportSection[];
  signatures: { role: string; name: string; date: string; status: 'Approuvé' | 'En attente' | 'Validé sous réserve' }[];
}

export const defaultReportTemplates: Record<string, ReportTemplate> = {
  'doc-etude-opportunite': {
    id: 'doc-etude-opportunite',
    name: "Étude d'Opportunité (Business Case)",
    category: 'Avant-Projet',
    code: 'DOC-OPP-01',
    title: "Dossier d'Étude d'Opportunité & Faisabilité",
    subtitle: "Évaluation stratégique, analyse de rentabilité et cadrage initial du projet",
    version: '1.0',
    date: new Date().toISOString().split('T')[0],
    author: 'Chef de Projet',
    approver: 'Comité de Direction',
    status: 'Validé',
    summary: "Ce document présente l'analyse d'opportunité du projet, détaillant les bénéfices qualitatifs et quantitatifs attendus, l'alignement avec les axes stratégiques de l'organisation ainsi que les scénarios de faisabilité économique et technique.",
    sections: [
      {
        id: 's1',
        title: '1. Contexte & Problématique Initiale',
        content: "L'organisation fait face à une nécessité de modernisation de ses processus opérationnels. Les solutions actuelles atteignent leurs limites capacitaires et engendrent des coûts récurrents élevés ainsi qu'un risque opérationnel accru. La mise en place d'un dispositif rénové permettra de sécuriser les opérations et d'améliorer la productivité.",
        type: 'text'
      },
      {
        id: 's2',
        title: '2. Objectifs Stratégiques & Bénéfices Attendus',
        content: "• Réduction du temps de traitement des dossiers de 35% d'ici 12 mois.\n• Amélioration de la traçabilité et conformité réglementaire à 100%.\n• Réduction des coûts d'exploitation annuels estimée à 45 000 €.\n• Augmentation du taux de satisfaction des parties prenantes internes et externes.",
        type: 'text'
      },
      {
        id: 's3',
        title: '3. Analyse Comparative des Scénarios',
        content: 'Évaluation des 3 scénarios envisagés pour répondre au besoin :',
        type: 'table',
        tableHeaders: ['Scénario', 'Coût estimé', 'Délai prévisionnel', 'Risque', 'Recommandation'],
        tableRows: [
          ['Statut Quo (Maintien de l\'existant)', '0 € initial', 'Immédiat', 'Très Élevé (Obsolescence)', 'Rejeté'],
          ['Développement Interne / Spécifique', '120 000 €', '10 mois', 'Modéré', 'Envisageable'],
          ['Déploiement Solution Intégrée Standard', '75 000 €', '5 mois', 'Faible', 'Retenu (Option recommandée)']
        ]
      },
      {
        id: 's4',
        title: '4. Analyse Économique & Retour sur Investissement (ROI)',
        content: "• Coût total estimé du projet (CAPEX + OPEX An 1) : 85 000 €\n• Gains bruts cumulés sur 3 ans : 160 000 €\n• Délai de récupération du capital investi (Payback) : 18 mois\n• Taux de Rendement Interne (TRI) : 24%",
        type: 'text'
      }
    ],
    signatures: [
      { role: 'Sponsor du Projet', name: 'Directeur Général / Sponsor', date: new Date().toISOString().split('T')[0], status: 'Approuvé' },
      { role: 'Chef de Projet', name: 'Responsable Projet', date: new Date().toISOString().split('T')[0], status: 'Approuvé' }
    ]
  },

  'doc-charte-projet': {
    id: 'doc-charte-projet',
    name: 'Charte de Projet',
    category: 'Initialisation',
    code: 'CHARTE-PRJ-01',
    title: 'Charte Fondatrice de Projet',
    subtitle: 'Mandat officiel, gouvernance, périmètre et engagements des parties prenantes',
    version: '1.0',
    date: new Date().toISOString().split('T')[0],
    author: 'Chef de Projet',
    approver: 'Comité de Pilotage (COPIL)',
    status: 'Validé',
    summary: "La Charte de Projet confère au Chef de Projet l'autorité formelle nécessaire pour mobiliser les ressources humaines, matérielles et financières allouées. Elle définit le cadre de gouvernance, les jalons macro et les limites de périmètre.",
    sections: [
      {
        id: 's1',
        title: '1. Justification & Raison d\'Être du Projet',
        content: "Conformément aux orientations stratégiques validées lors du séminaire de direction, ce projet vise à poser les fondations techniques et organisationnelles pour l'exercice à venir, assurant l'alignement entre les exigences métiers et les capacités de livraison.",
        type: 'text'
      },
      {
        id: 's2',
        title: '2. Périmètre (In Scope / Out of Scope)',
        content: 'Détail des inclusions et exclusions formelles :',
        type: 'table',
        tableHeaders: ['Élément', 'Statut Périmètre', 'Justification / Impact'],
        tableRows: [
          ['Paramétrage et intégration du socle fonctionnel', 'In Scope (Inclus)', 'Livrable cœur de la mission'],
          ['Migration des données historiques (3 dernières années)', 'In Scope (Inclus)', 'Indispensable pour la continuité opérationnelle'],
          ['Développement d\'interfaces sur mesure non standardisées', 'Out of Scope (Exclu)', 'Reporté en phase ultérieure (Phase 2)'],
          ['Formation des utilisateurs finaux et supports pédagogiques', 'In Scope (Inclus)', 'Garantit l\'adoption de la solution']
        ]
      },
      {
        id: 's3',
        title: '3. Macro-Planning & Grands Jalons Contractuels',
        content: 'Calendrier des jalons décisionnels :',
        type: 'table',
        tableHeaders: ['Jalon', 'Description', 'Date Cible', 'Livrable Majeur associé'],
        tableRows: [
          ['J0 - Lancement', 'Validation de la charte et kick-off', 'Semaine 1', 'Charte signée, registre prenantes'],
          ['J1 - Cadrage validé', 'Spécifications et cahier des charges approuvés', 'Semaine 6', 'Cahier des charges signé'],
          ['J2 - Fin de réalisation', 'Livraison de la solution configurée', 'Semaine 14', 'Version candidate pour recette'],
          ['J3 - Recette Utilisateur', 'Approbation sans réserve bloquante', 'Semaine 18', 'PV de recette avec réserve levée'],
          ['J4 - Déploiement & Clôture', 'Mise en production et bilan', 'Semaine 20', 'PV de réception, bilan final']
        ]
      },
      {
        id: 's4',
        title: '4. Gouvernance & Instance Décisionnelle',
        content: "• Comité de Pilotage (COPIL) : Réunion mensuelle pour arbitrages budgétaires et modifications de périmètre.\n• Comité de Projet (COPROJ) : Réunion hebdomadaire opérationnelle d'avancement des chantiers.\n• Revue des Risques & Actions : Intégrée aux points d'avancement hebdos.",
        type: 'text'
      }
    ],
    signatures: [
      { role: 'Sponsor Exécutif', name: 'Direction Métier', date: new Date().toISOString().split('T')[0], status: 'Approuvé' },
      { role: 'Directeur de Projet', name: 'Direction des Projets', date: new Date().toISOString().split('T')[0], status: 'Approuvé' },
      { role: 'Chef de Projet Mandaté', name: 'Responsable Opérationnel', date: new Date().toISOString().split('T')[0], status: 'Approuvé' }
    ]
  },

  'doc-expression-besoins': {
    id: 'doc-expression-besoins',
    name: 'Expression des Besoins',
    category: 'Planification',
    code: 'DOC-BES-01',
    title: 'Dossier d\'Expression des Besoins Métiers',
    subtitle: 'Recueil exhaustif des exigences fonctionnelles, organisationnelles et techniques',
    version: '1.0',
    date: new Date().toISOString().split('T')[0],
    author: 'Responsable Métier / PO',
    approver: 'Chef de Projet',
    status: 'Validé',
    summary: "Ce document formalise les attentes détaillées des directions opérationnelles et parties prenantes. Il sert de socle pour la rédaction du Cahier des Charges Fonctionnel et l'arbitrage des priorités.",
    sections: [
      {
        id: 's1',
        title: '1. Contexte Opérationnel & Utilisateurs Cibles',
        content: "Les profils utilisateurs concernés regroupent les gestionnaires de projet, les coordinateurs techniques, la direction des opérations et les équipes d'audit qualité.",
        type: 'text'
      },
      {
        id: 's2',
        title: '2. Grille des Besoins Fonctionnels Priorisés (MoSCoW)',
        content: 'Exigences recensées auprès des équipes métiers :',
        type: 'table',
        tableHeaders: ['Réf. Besoin', 'Description de l\'Exigence', 'Priorité (MoSCoW)', 'Critère d\'Acceptation'],
        tableRows: [
          ['BES-01', 'Génération automatisée de rapports imprimables et exportables en PDF', 'Must Have (Indispensable)', 'Export PDF instantané avec logo entreprise'],
          ['BES-02', 'Saisie simplifiée avec formulaires pré-remplis et contrôles de saisie', 'Must Have (Indispensable)', 'Contrôle cohérence sur champs obligatoires'],
          ['BES-03', 'Visualiseur de supports de formation interactifs par diapositives', 'Should Have (Important)', 'Navigation fluide précédent/suivant et synthèse'],
          ['BES-04', 'Conservation locale sécurisée des données sans dépendance serveur externe', 'Must Have (Indispensable)', 'Données sauvegardées en mode local / offline'],
          ['BES-05', 'Personnalisation des modèles de comptes-rendus et livrables', 'Could Have (Souhaitable)', 'Possibilité d\'ajouter des sections sur mesure']
        ]
      }
    ],
    signatures: [
      { role: 'Représentant Métier', name: 'Responsable Utilisateurs', date: new Date().toISOString().split('T')[0], status: 'Approuvé' },
      { role: 'Chef de Projet', name: 'Coordinateur Projet', date: new Date().toISOString().split('T')[0], status: 'Approuvé' }
    ]
  },

  'doc-cahier-des-charges': {
    id: 'doc-cahier-des-charges',
    name: 'Cahier des Charges Fonctionnel',
    category: 'Planification',
    code: 'CDCF-2026-01',
    title: 'Cahier des Charges Fonctionnel (CdCF)',
    subtitle: 'Spécifications fonctionnelles, contraintes techniques, critères d\'ergonomie et livrables',
    version: '1.1',
    date: new Date().toISOString().split('T')[0],
    author: 'Chef de Projet & Architecte',
    approver: 'Comité de Projet',
    status: 'Validé',
    summary: "Ce document traduit l'expression des besoins en spécifications fonctionnelles précises. Il constitue la base contractuelle des engagements de réalisation pour l'équipe de développement et de déploiement.",
    sections: [
      {
        id: 's1',
        title: '1. Architecture & Contraintes d\'Intégration',
        content: "• Application client léger Desktop haute performance construite avec technologies modernes.\n• Zéro latence, réactivité instantanée pour toutes les interactions de saisie et de navigation.\n• Stockage local persistant garantissant la confidentialité absolue des données de gestion de projet.",
        type: 'text'
      },
      {
        id: 's2',
        title: '2. Matrice des Spécifications Fonctionnelles Détaillées',
        content: 'Matrice de couverture fonctionnelle :',
        type: 'table',
        tableHeaders: ['Module', 'Fonctionnalité', 'Niveau d\'Exigence', 'Règle de Gestion'],
        tableRows: [
          ['Rapports & Édition', 'Formulaire interactif avec pré-visualisation', 'Critique', 'Mise à jour en temps réel de l\'aperçu A4'],
          ['Export PDF', 'Impression et export vectoriel au format A4', 'Critique', 'Bordures propres, en-tête avec logo société, zéro coupure gênante'],
          ['Formations', 'Supports de formation par diapositives', 'Élevé', 'Lecteur interactif complet avec fiches méthodologiques'],
          ['Logo Société', 'Importation et mémorisation du logo d\'entreprise', 'Critique', 'Conservation automatique au format base64']
        ]
      }
    ],
    signatures: [
      { role: 'Chef de Projet', name: 'Chef de Projet', date: new Date().toISOString().split('T')[0], status: 'Approuvé' },
      { role: 'Architecte Solution', name: 'Responsable Technique', date: new Date().toISOString().split('T')[0], status: 'Approuvé' }
    ]
  },

  'doc-poc': {
    id: 'doc-poc',
    name: 'Dossier Proof of Concept (POC)',
    category: 'Planification',
    code: 'POC-VAL-01',
    title: 'Dossier de Validation Proof of Concept (POC)',
    subtitle: 'Rapport d\'expérimentation, tests de faisabilité technique et validation d\'hypothèses',
    version: '1.0',
    date: new Date().toISOString().split('T')[0],
    author: 'Équipe R&D / Projets',
    approver: 'Direction Technique',
    status: 'Validé',
    summary: "Rapport de synthèse du POC démontrant la faisabilité de la solution retenue, les performances mesurées et les retours d'expérience avant engagement du budget de déploiement général.",
    sections: [
      {
        id: 's1',
        title: '1. Objectifs & Hypothèses du POC',
        content: "L'objectif était de vérifier la viabilité technique d'une application Desktop moderne combinant saisie de formulaires, génération de PDF haute qualité et chargement instantané de supports de formation sans recours à une connexion internet constante.",
        type: 'text'
      },
      {
        id: 's2',
        title: '2. Résultats des Tests de Performance',
        content: 'Résultats quantitatifs obtenus lors du banc d\'essai :',
        type: 'table',
        tableHeaders: ['Indicateur mesuré', 'Cible fixée', 'Résultat obtenu', 'Verdict'],
        tableRows: [
          ['Temps d\'ouverture de l\'application', '< 2.0 s', '0.6 s', 'Succès remarquable'],
          ['Génération et aperçu du rapport PDF', '< 1.5 s', '0.2 s instantané', 'Succès remarquable'],
          ['Taille mémoire consommée', '< 150 Mo', '68 Mo', 'Conforme aux attentes'],
          ['Taux de réussite export avec logo', '100%', '100% sans bavure', 'Validé']
        ]
      }
    ],
    signatures: [
      { role: 'Responsable POC', name: 'Expert Technique', date: new Date().toISOString().split('T')[0], status: 'Approuvé' }
    ]
  },

  'doc-plan-communication': {
    id: 'doc-plan-communication',
    name: 'Plan de Communication',
    category: 'Planification',
    code: 'PLAN-COM-01',
    title: 'Plan de Communication du Projet',
    subtitle: 'Stratégie de communication, messages clés, canaux et calendrier de diffusion',
    version: '1.0',
    date: new Date().toISOString().split('T')[0],
    author: 'Chef de Projet',
    approver: 'Sponsor & COPIL',
    status: 'Validé',
    summary: "Ce plan identifie les cibles, les objectifs de communication, les supports adaptés et le rythme des réunions et synthèses pour assurer l'adhésion totale des acteurs au projet.",
    sections: [
      {
        id: 's1',
        title: '1. Matrice des Canaux & Actions de Communication',
        content: 'Planning et canaux de communication institutionnelle et projet :',
        type: 'table',
        tableHeaders: ['Cible', 'Message Clé', 'Canal / Support', 'Fréquence', 'Responsable'],
        tableRows: [
          ['Direction Générale & Sponsor', 'Arbitrages, santé globale du projet, jalons stratégiques', 'Réunion COPIL & Synthèse Décisionnelle', 'Mensuelle', 'Chef de Projet'],
          ['Équipes Opérationnelles', 'Avancement des tâches, points de blocage, priorités', 'Comité de Projet (COPROJ)', 'Hebdomadaire', 'Chef de Projet'],
          ['Utilisateurs Finaux', 'Découverte de la solution, formations, planning déploiement', 'Supports de formation interactifs & Ateliers', 'Selon jalons clés', 'Équipe Projet'],
          ['Toutes parties prenantes', 'Météo de synthèse, faits marquants, jalons franchis', 'Rapport Flash Projet', 'Hebdomadaire', 'Chef de Projet']
        ]
      }
    ],
    signatures: [
      { role: 'Chef de Projet', name: 'Chef de Projet', date: new Date().toISOString().split('T')[0], status: 'Approuvé' }
    ]
  },

  'doc-plan-projet': {
    id: 'doc-plan-projet',
    name: 'Plan de Management de Projet (PMP)',
    category: 'Planification',
    code: 'PMP-SYNTH-01',
    title: 'Plan de Management de Projet (PMP)',
    subtitle: 'Document cadre de référence régissant la gouvernance, l\'exécution et le contrôle',
    version: '1.2',
    date: new Date().toISOString().split('T')[0],
    author: 'Directeur de Projet',
    approver: 'Comité de Direction',
    status: 'Validé',
    summary: "Le Plan de Management de Projet (PMP) formalise l'ensemble des règles de pilotage : gestion du périmètre, gestion des délais, gestion des coûts, gestion des risques et management de la qualité.",
    sections: [
      {
        id: 's1',
        title: '1. Cadre de Pilotage & Méthodologie',
        content: "Le projet applique les meilleures pratiques de gestion de projet structurée en 6 phases séquencées avec portes de passage (Gate Reviews) pour chaque étape clé.",
        type: 'text'
      },
      {
        id: 's2',
        title: '2. Dispositif de Pilotage des Risques & Alertes',
        content: "• Registre des risques actualisé à chaque fin de semaine.\n• Escalade immédiate vers le Sponsor pour tout risque dont l'impact dépasse 10 jours de dérive ou 10% du budget alloué.\n• Processus formel de demande de changement (Workflow de modification).",
        type: 'text'
      }
    ],
    signatures: [
      { role: 'Directeur de Projet', name: 'Directeur de Projet', date: new Date().toISOString().split('T')[0], status: 'Approuvé' },
      { role: 'Sponsor', name: 'Directeur Métier', date: new Date().toISOString().split('T')[0], status: 'Approuvé' }
    ]
  },

  'doc-plan-qualite': {
    id: 'doc-plan-qualite',
    name: "Plan d'Assurance Qualité (PAQ)",
    category: 'Planification',
    code: 'PAQ-2026-01',
    title: "Plan d'Assurance Qualité (PAQ)",
    subtitle: 'Standards, métriques de conformité, revues de pairs et critères d\'acceptation',
    version: '1.0',
    date: new Date().toISOString().split('T')[0],
    author: 'Responsable Qualité Projet',
    approver: 'Directeur de Projet',
    status: 'Validé',
    summary: "Définit les exigences et les contrôles qualité appliqués aux livrables, aux processus de travail et aux validations intermédiaires afin de garantir l'excellence opérationnelle.",
    sections: [
      {
        id: 's1',
        title: '1. Métriques et Seuils de Tolérance Qualité',
        content: 'Indicateurs de conformité applicables :',
        type: 'table',
        tableHeaders: ['Critère Qualité', 'Métrique / Cible', 'Fréquence de Contrôle', 'Action en cas d\'écart'],
        tableRows: [
          ['Exhaustivité des livrables', '100% des attendus documentaires', 'Fin de phase', 'Revue de complétude obligatoire'],
          ['Taux d\'anomalies bloquantes', '0 anomalie bloquante en recette', 'Continu pendant la phase test', 'Gel du passage au jalon suivant'],
          ['Satisfaction utilisateurs', 'Score >= 85/100', 'Post-formation', 'Plan d\'action correctif sous 5 jours']
        ]
      }
    ],
    signatures: [
      { role: 'Responsable Qualité', name: 'Auditeur Qualité', date: new Date().toISOString().split('T')[0], status: 'Approuvé' }
    ]
  },

  'doc-estimation-budget': {
    id: 'doc-estimation-budget',
    name: "Rapport d'Estimation Budgétaire",
    category: 'Planification',
    code: 'EST-BUD-01',
    title: "Rapport d'Estimation Budgétaire & Coûts",
    subtitle: 'Chiffrage détaillé par poste de dépenses, ressources et provisions pour risques',
    version: '1.0',
    date: new Date().toISOString().split('T')[0],
    author: 'Contrôleur de Gestion & Chef de Projet',
    approver: 'Direction Financière',
    status: 'Validé',
    summary: "Présentation détaillée du budget prévisionnel alloué au projet, ventilé par lots de travaux, charges internes, prestataires externes et marge pour aléas.",
    sections: [
      {
        id: 's1',
        title: '1. Ventilation Budgétaire par Poste',
        content: 'Répartition analytique des investissements prévus :',
        type: 'table',
        tableHeaders: ['Poste de Dépense', 'Budget Alloué (€)', 'Part relative (%)', 'Commentaires'],
        tableRows: [
          ['Conception & Architecture', '15 000 €', '17.6%', 'Spécifications et cadrage fonctionnel'],
          ['Réalisation & Déploiement', '38 000 €', '44.7%', 'Développement, intégration et paramétrage'],
          ['Recette & Tests de conformité', '12 000 €', '14.1%', 'Cahiers de tests et vérifications'],
          ['Accompagnement au Changement & Formation', '10 000 €', '11.8%', 'Création des supports et sessions'],
          ['Provisions pour aléas & risques (Contingence)', '10 000 €', '11.8%', 'Réserve sous arbitrage du Chef de projet'],
          ['TOTAL GÉNÉRAL', '85 000 €', '100.0%', 'Enveloppe globale validée']
        ]
      }
    ],
    signatures: [
      { role: 'Directeur Financier', name: 'Direction Financière', date: new Date().toISOString().split('T')[0], status: 'Approuvé' },
      { role: 'Chef de Projet', name: 'Chef de Projet', date: new Date().toISOString().split('T')[0], status: 'Approuvé' }
    ]
  },

  'doc-cr-reunion': {
    id: 'doc-cr-reunion',
    name: 'Compte-Rendu Synthétique de Réunion',
    category: 'Exécution',
    code: 'CR-SYNTH-24',
    title: 'Compte-Rendu Synthétique de Réunion',
    subtitle: 'Suivi des échanges, relevé des décisions et plan d\'action opérationnel',
    version: '1.0',
    date: new Date().toISOString().split('T')[0],
    author: 'Chef de Projet',
    approver: 'Participants à la réunion',
    status: 'Diffusé',
    summary: "Ce compte-rendu synthétise les points majeurs abordés lors de la séance de travail, les arbitrages actés ainsi que les engagements pris par chaque intervenant avec date d'échéance.",
    sections: [
      {
        id: 's1',
        title: '1. Informations Générales de la Séance',
        content: "• Type de réunion : Comité de Projet Opérationnel (Point d'Avancement)\n• Date & Heure : " + new Date().toLocaleDateString('fr-FR') + " de 10h00 à 11h30\n• Lieu / Modalité : Salle de réunion principale & Visioconférence\n• Ordre du jour : Revue des livrables de la semaine, points de blocage, préparation des tests.",
        type: 'text'
      },
      {
        id: 's2',
        title: '2. Relevé des Décisions Actées',
        content: 'Décisions fermes prises en séance :',
        type: 'table',
        tableHeaders: ['Réf. Décision', 'Intitulé de la Décision', 'Porteur', 'Impact'],
        tableRows: [
          ['DEC-01', 'Validation de l\'intégration des rapports officiels dans le menu de pilotage', 'Chef de Projet', 'Permet le passage immédiat en phase de test'],
          ['DEC-02', 'Activation de l\'importateur de logo d\'entreprise pour tous les exports PDF', 'Responsable Produit', 'Améliore la conformité corporate des documents'],
          ['DEC-03', 'Maintien de la date de livraison finale au calendrier prévu sans dérive', 'Sponsor', 'Confirme le respect du jalon contractuel J4']
        ]
      },
      {
        id: 's3',
        title: '3. Plan d\'Actions & Engagements',
        content: 'Tâches assignées lors de la réunion :',
        type: 'table',
        tableHeaders: ['Action', 'Description', 'Responsable', 'Échéance', 'Statut'],
        tableRows: [
          ['ACT-01', 'Finaliser les tests d\'impression et de génération PDF', 'Équipe Technique', 'J+2', 'En cours'],
          ['ACT-02', 'Intégrer les présentations de formation interactives', 'Équipe Pédagogique', 'J+3', 'En cours'],
          ['ACT-03', 'Diffuser le compte-rendu à tous les membres du comité', 'Chef de Projet', 'J+1', 'Réalisé']
        ]
      }
    ],
    signatures: [
      { role: 'Rédacteur', name: 'Chef de Projet', date: new Date().toISOString().split('T')[0], status: 'Approuvé' },
      { role: 'Validation séance', name: 'Président de séance', date: new Date().toISOString().split('T')[0], status: 'Approuvé' }
    ]
  },

  'doc-minutes-copil': {
    id: 'doc-minutes-copil',
    name: 'Procès-Verbal Réunion COPIL',
    category: 'Exécution',
    code: 'PV-COPIL-2026',
    title: 'Procès-Verbal / Minutes Réunion Comité de Pilotage',
    subtitle: 'Revue de gouvernance, arbitrage stratégique et validation des passages de jalons',
    version: '1.0',
    date: new Date().toISOString().split('T')[0],
    author: 'Chef de Projet',
    approver: 'Membres du Comité de Pilotage',
    status: 'Validé',
    summary: "Relevé officiel des délibérations et arbitrages du Comité de Pilotage (COPIL), garantissant la traçabilité des décisions managériales et l'autorisation de franchissement des portes d'étapes.",
    sections: [
      {
        id: 's1',
        title: '1. Composition du Comité & Présences',
        content: "• Président de séance : Directeur Général / Sponsor\n• Membres présents : Directeur Financier, Directeur Métier, Responsable Qualité, Chef de Projet\n• Objet : Revue trimestrielle des engagements et validation du jalon d'exécution.",
        type: 'text'
      },
      {
        id: 's2',
        title: '2. Arbitrages & Décisions Stratégiques du COPIL',
        content: 'Arbitrages formellement votés en séance :',
        type: 'table',
        tableHeaders: ['N° Arbitrage', 'Objet de la Décision', 'Décision COPIL', 'Conséquence Opérationnelle'],
        tableRows: [
          ['ARB-01', 'Acceptation du franchissement du Jalon J2 (Fin de développement)', 'APPROUVÉ À L\'UNANIMITÉ', 'Autorisation d\'engager la campagne de recette'],
          ['ARB-02', 'Affectation d\'une ressource complémentaire pour l\'assistance recette', 'APPROUVÉ', 'Mobilisation d\'un profil métier pendant 3 semaines'],
          ['ARB-03', 'Maintien de l\'enveloppe budgétaire globale sans dépassement', 'CONFIRMÉ', 'Aucun avenant financier nécessaire']
        ]
      }
    ],
    signatures: [
      { role: 'Président du COPIL', name: 'Direction Générale', date: new Date().toISOString().split('T')[0], status: 'Approuvé' },
      { role: 'Secrétaire de séance', name: 'Chef de Projet', date: new Date().toISOString().split('T')[0], status: 'Approuvé' }
    ]
  },

  'doc-strategie-tests': {
    id: 'doc-strategie-tests',
    name: 'Stratégie de Recette et Tests',
    category: 'Exécution',
    code: 'TEST-STRAT-01',
    title: 'Stratégie de Qualification, Tests et Recette',
    subtitle: 'Organisation des campagnes de tests, environnements, critères d\'acceptation et seuils de qualité',
    version: '1.0',
    date: new Date().toISOString().split('T')[0],
    author: 'Responsable Recette & Assurance Qualité',
    approver: 'Chef de Projet',
    status: 'Validé',
    summary: "Ce document décrit la méthodologie globale de test applicable au projet, définissant les types de tests (unitaires, intégration, bout-en-bout, performance, recette utilisateur) et les critères de sortie de recette.",
    sections: [
      {
        id: 's1',
        title: '1. Périmètre & Niveaux de Tests',
        content: "• Tests Unitaires & Composants : Vérification de la conformité du code et absence de régressions.\n• Tests d'Intégration & Données : Validation de la persistance locale et sauvegarde des états.\n• Tests de Rendu & Impression : Vérification du calibrage A4 et de l'export vectoriel PDF.\n• Recette Utilisateur (UAT) : Scénarios réels exécutés par les utilisateurs métiers.",
        type: 'text'
      },
      {
        id: 's2',
        title: '2. Critères d\'Entrée et de Sortie de Recette',
        content: 'Conditions strictes de validation :',
        type: 'table',
        tableHeaders: ['Étape', 'Critère Exigé', 'Méthode de Mesure', 'Responsable Validation'],
        tableRows: [
          ['Entrée en Recette', '100% des cas de tests documentés et environnement stable', 'Audit de préparation', 'Responsable Tests'],
          ['Sortie de Recette', '0 anomalie bloquante, moins de 2 anomalies mineures contournables', 'Bilan des anomalies', 'Chef de Projet & Métier'],
          ['Validation Finale', 'Signature sans réserve du Procès-Verbal de Recette', 'Document officiel signé', 'Comité de Recette']
        ]
      }
    ],
    signatures: [
      { role: 'Responsable des Tests', name: 'Lead QA', date: new Date().toISOString().split('T')[0], status: 'Approuvé' },
      { role: 'Chef de Projet', name: 'Chef de Projet', date: new Date().toISOString().split('T')[0], status: 'Approuvé' }
    ]
  },

  'doc-cahier-recette': {
    id: 'doc-cahier-recette',
    name: 'Cahier de Recette Utilisateur',
    category: 'Exécution',
    code: 'RECETTE-UAT-01',
    title: 'Cahier de Recette & Validation Utilisateur (UAT)',
    subtitle: 'Fiches de tests opérationnels, résultats constatés et traçabilité des validations',
    version: '1.0',
    date: new Date().toISOString().split('T')[0],
    author: 'Équipe Recette Métier',
    approver: 'Responsable Produit',
    status: 'Validé',
    summary: "Ce cahier rassemble l'ensemble des cas d'usage métiers testés, consignant pour chaque scénario le résultat attendu, le résultat obtenu et le statut de validation.",
    sections: [
      {
        id: 's1',
        title: '1. Grille d\'Exécution des Scénarios de Test',
        content: 'Résultat d\'exécution des cas de recette fonctionnelle :',
        type: 'table',
        tableHeaders: ['ID Test', 'Scénario Opérationnel Testé', 'Résultat Attendu', 'Résultat Constaté', 'Statut'],
        tableRows: [
          ['TEST-01', 'Saisie d\'un rapport avec données personnalisées', 'Mise à jour immédiate de l\'aperçu', 'Aperçu fluide et temps réel', 'CONFORME'],
          ['TEST-02', 'Import d\'un logo société au format PNG/JPG', 'Le logo s\'affiche dans l\'en-tête du document', 'Logo mémorisé et visible au pixel près', 'CONFORME'],
          ['TEST-03', 'Impression et export PDF du document officiel', 'Mise en page A4 impeccable sans débordement', 'Document A4 officiel parfait', 'CONFORME'],
          ['TEST-04', 'Consultation des diapositives de formation interactive', 'Navigation diapositive par diapositive opérationnelle', 'Slide viewer interactif rapide', 'CONFORME'],
          ['TEST-05', 'Suppression et réinitialisation de formulaire', 'Action immédiate avec notification de confirmation', 'Fonctionnement nominal', 'CONFORME']
        ]
      }
    ],
    signatures: [
      { role: 'Testeur Métier', name: 'Référent Opérationnel', date: new Date().toISOString().split('T')[0], status: 'Approuvé' },
      { role: 'Responsable Recette', name: 'Coordinateur Qualité', date: new Date().toISOString().split('T')[0], status: 'Approuvé' }
    ]
  },

  'doc-pv-reception': {
    id: 'doc-pv-reception',
    name: 'PV de Réception des Livrables',
    category: 'Clôture',
    code: 'PV-REC-LIV-33',
    title: 'Procès-Verbal de Réception des Livrables de Projet',
    subtitle: 'Constat officiel de conformité, acceptation des réalisations et clôture formelle',
    version: '1.0',
    date: new Date().toISOString().split('T')[0],
    author: 'Chef de Projet',
    approver: 'Comité de Clôture & Client',
    status: 'Validé',
    summary: "Le présent Procès-Verbal atteste de la livraison effective et de l'acceptation sans réserve majeure des livrables contractuels du projet par le commanditaire.",
    sections: [
      {
        id: 's1',
        title: '1. Constat de Livraison des Produits du Projet',
        content: 'Récapitulatif de l\'état d\'acceptation des livrables majeurs :',
        type: 'table',
        tableHeaders: ['Réf. Livrable', 'Désignation du Livrable', 'Version', 'Statut de Réception', 'Observations'],
        tableRows: [
          ['LIV-01', 'Plateforme FAAB PROJET Desktop opérationnelle', 'v3.0', 'Accepté sans réserve', 'Déployée et validée'],
          ['LIV-02', 'Générateur de rapports officiels & exports PDF', 'v1.0', 'Accepté sans réserve', 'Logo entreprise intégré avec succès'],
          ['LIV-03', 'Ensemble des modules de formation interactifs (6 cursus)', 'v1.0', 'Accepté sans réserve', 'Supports pédagogiques complets'],
          ['LIV-04', 'Documentation d\'exploitation et guides utilisateurs', 'v1.0', 'Accepté sans réserve', 'Conforme aux exigences']
        ]
      },
      {
        id: 's2',
        title: '2. Clôture des Engagements Contractuels & Garantie',
        content: "Par la signature du présent acte, le client commanditaire déclare recevoir l'ensemble des éléments susmentionnés. La période de garantie opérationnelle et de support démarre à compter de ce jour pour une durée de 3 mois.",
        type: 'text'
      }
    ],
    signatures: [
      { role: 'Le Commanditaire / Client', name: 'Directeur Métier', date: new Date().toISOString().split('T')[0], status: 'Approuvé' },
      { role: 'Le Chef de Projet', name: 'Responsable de Projet', date: new Date().toISOString().split('T')[0], status: 'Approuvé' },
      { role: 'Direction Qualité', name: 'Auditeur Projet', date: new Date().toISOString().split('T')[0], status: 'Approuvé' }
    ]
  }
};

interface ProjectReportGeneratorProps {
  projectId: string;
  fileId?: string;
  onClose?: () => void;
}

export const ProjectReportGenerator: React.FC<ProjectReportGeneratorProps> = ({
  projectId,
  fileId = 'doc-charte-projet'
}) => {
  // Select active template
  const [selectedTemplateId, setSelectedTemplateId] = useState<string>(
    defaultReportTemplates[fileId] ? fileId : 'doc-charte-projet'
  );

  // Logo state
  const [companyLogo, setCompanyLogo] = useState<string | null>(null);
  const [viewMode, setViewMode] = useState<'split' | 'form' | 'preview'>('split');
  const [reportData, setReportData] = useState<ReportTemplate>(
    defaultReportTemplates[selectedTemplateId] || defaultReportTemplates['doc-charte-projet']
  );
  const [savedNotification, setSavedNotification] = useState<string | null>(null);

  const fileInputRef = useRef<HTMLInputElement>(null);

  // Load saved logo from localStorage
  useEffect(() => {
    const savedGlobalLogo = localStorage.getItem('faab-enterprise-logo');
    const savedProjectLogo = localStorage.getItem(`faab-logo-${projectId}`);
    if (savedProjectLogo) {
      setCompanyLogo(savedProjectLogo);
    } else if (savedGlobalLogo) {
      setCompanyLogo(savedGlobalLogo);
    }
  }, [projectId]);

  // Load saved report data or fallback to template
  useEffect(() => {
    const targetId = defaultReportTemplates[fileId] ? fileId : selectedTemplateId;
    setSelectedTemplateId(targetId);

    const storageKey = `faab-report-data-${projectId}-${targetId}`;
    const saved = localStorage.getItem(storageKey);
    if (saved) {
      try {
        setReportData(JSON.parse(saved));
      } catch (e) {
        setReportData(defaultReportTemplates[targetId] || defaultReportTemplates['doc-charte-projet']);
      }
    } else {
      setReportData(defaultReportTemplates[targetId] || defaultReportTemplates['doc-charte-projet']);
    }
  }, [fileId, selectedTemplateId, projectId]);

  // Handle Logo Upload
  const handleLogoUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (file) {
      if (file.size > 2 * 1024 * 1024) {
        alert("L'image dépasse 2 Mo. Veuillez sélectionner un fichier plus léger.");
        return;
      }
      const reader = new FileReader();
      reader.onload = (event) => {
        const base64 = event.target?.result as string;
        setCompanyLogo(base64);
        localStorage.setItem(`faab-logo-${projectId}`, base64);
        localStorage.setItem('faab-enterprise-logo', base64);
        triggerNotification("Logo de l'entreprise enregistré avec succès !");
      };
      reader.readAsDataURL(file);
    }
  };

  const handleRemoveLogo = () => {
    setCompanyLogo(null);
    localStorage.removeItem(`faab-logo-${projectId}`);
    localStorage.removeItem('faab-enterprise-logo');
    triggerNotification("Logo supprimé.");
  };

  const triggerNotification = (msg: string) => {
    setSavedNotification(msg);
    setTimeout(() => setSavedNotification(null), 3000);
  };

  // Save changes to localStorage
  const handleSave = () => {
    const storageKey = `faab-report-data-${projectId}-${selectedTemplateId}`;
    localStorage.setItem(storageKey, JSON.stringify(reportData));
    triggerNotification("Rapport sauvegardé avec succès !");
  };

  // Reset to default
  const handleReset = () => {
    if (window.confirm("Voulez-vous réinitialiser ce document avec son contenu d'origine ?")) {
      const initial = defaultReportTemplates[selectedTemplateId];
      if (initial) {
        setReportData(JSON.parse(JSON.stringify(initial)));
        const storageKey = `faab-report-data-${projectId}-${selectedTemplateId}`;
        localStorage.removeItem(storageKey);
        triggerNotification("Modèle réinitialisé.");
      }
    }
  };

  // Print & PDF Export
  const handlePrint = () => {
    window.print();
  };

  // Update Section Text
  const updateSectionContent = (idx: number, newContent: string) => {
    const updated = { ...reportData };
    updated.sections[idx].content = newContent;
    setReportData(updated);
  };

  // Update Table Cell
  const updateTableCell = (sIdx: number, rIdx: number, cIdx: number, val: string) => {
    const updated = { ...reportData };
    if (updated.sections[sIdx].tableRows) {
      updated.sections[sIdx].tableRows![rIdx][cIdx] = val;
      setReportData(updated);
    }
  };

  // Add Table Row
  const addTableRow = (sIdx: number) => {
    const updated = { ...reportData };
    const headers = updated.sections[sIdx].tableHeaders || [];
    const newRow = headers.map(() => 'Nouveau...');
    if (!updated.sections[sIdx].tableRows) {
      updated.sections[sIdx].tableRows = [];
    }
    updated.sections[sIdx].tableRows!.push(newRow);
    setReportData(updated);
  };

  // Delete Table Row
  const deleteTableRow = (sIdx: number, rIdx: number) => {
    const updated = { ...reportData };
    if (updated.sections[sIdx].tableRows && updated.sections[sIdx].tableRows!.length > 1) {
      updated.sections[sIdx].tableRows!.splice(rIdx, 1);
      setReportData(updated);
    }
  };

  return (
    <div className="report-generator-container" style={{ display: 'flex', flexDirection: 'column', height: '100%', background: '#0f172a', color: '#f8fafc', overflow: 'hidden' }}>
      {/* Top Action Bar (Hidden on print) */}
      <div className="no-print" style={{
        padding: '0.8rem 1.5rem',
        background: 'rgba(15, 23, 42, 0.95)',
        borderBottom: '1px solid rgba(255, 255, 255, 0.1)',
        display: 'flex',
        alignItems: 'center',
        justifyContent: 'space-between',
        flexWrap: 'wrap',
        gap: '1rem',
        zIndex: 10
      }}>
        {/* Left: Template Selector */}
        <div style={{ display: 'flex', alignItems: 'center', gap: '0.8rem' }}>
          <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
            <span style={{ fontSize: '1.3rem' }}>📄</span>
            <div>
              <div style={{ fontSize: '0.75rem', color: '#94a3b8', textTransform: 'uppercase', fontWeight: 600 }}>Rapports Officiels & Documents PDF</div>
              <select
                value={selectedTemplateId}
                onChange={(e) => setSelectedTemplateId(e.target.value)}
                style={{
                  background: 'rgba(30, 41, 59, 0.9)',
                  border: '1px solid rgba(56, 189, 248, 0.4)',
                  color: '#38bdf8',
                  padding: '5px 12px',
                  borderRadius: '6px',
                  fontWeight: 700,
                  fontSize: '0.9rem',
                  cursor: 'pointer'
                }}
              >
                {Object.values(defaultReportTemplates).map(tmpl => (
                  <option key={tmpl.id} value={tmpl.id} style={{ background: '#0f172a', color: '#f8fafc' }}>
                    [{tmpl.category}] {tmpl.name} ({tmpl.code})
                  </option>
                ))}
              </select>
            </div>
          </div>
        </div>

        {/* Center: View Switcher */}
        <div style={{ display: 'flex', background: 'rgba(255, 255, 255, 0.06)', borderRadius: '8px', padding: '3px' }}>
          <button
            onClick={() => setViewMode('split')}
            style={{
              padding: '6px 14px',
              border: 'none',
              background: viewMode === 'split' ? '#0284c7' : 'transparent',
              color: '#fff',
              borderRadius: '6px',
              fontSize: '0.8rem',
              fontWeight: 600,
              cursor: 'pointer'
            }}
          >
            ⚖️ Vue Partagée
          </button>
          <button
            onClick={() => setViewMode('form')}
            style={{
              padding: '6px 14px',
              border: 'none',
              background: viewMode === 'form' ? '#0284c7' : 'transparent',
              color: '#fff',
              borderRadius: '6px',
              fontSize: '0.8rem',
              fontWeight: 600,
              cursor: 'pointer'
            }}
          >
            ✏️ Formulaire
          </button>
          <button
            onClick={() => setViewMode('preview')}
            style={{
              padding: '6px 14px',
              border: 'none',
              background: viewMode === 'preview' ? '#0284c7' : 'transparent',
              color: '#fff',
              borderRadius: '6px',
              fontSize: '0.8rem',
              fontWeight: 600,
              cursor: 'pointer'
            }}
          >
            👁️ Aperçu Document
          </button>
        </div>

        {/* Right: Actions (Logo, Save, Reset, Export PDF) */}
        <div style={{ display: 'flex', alignItems: 'center', gap: '0.8rem' }}>
          {/* Logo upload button */}
          <input
            type="file"
            ref={fileInputRef}
            onChange={handleLogoUpload}
            accept="image/*"
            style={{ display: 'none' }}
          />
          <button
            onClick={() => fileInputRef.current?.click()}
            style={{
              background: companyLogo ? 'rgba(16, 185, 129, 0.15)' : 'rgba(255, 255, 255, 0.08)',
              border: `1px solid ${companyLogo ? '#10b981' : 'rgba(255, 255, 255, 0.2)'}`,
              color: companyLogo ? '#10b981' : '#cbd5e1',
              padding: '6px 12px',
              borderRadius: '6px',
              fontSize: '0.8rem',
              fontWeight: 600,
              display: 'flex',
              alignItems: 'center',
              gap: '6px',
              cursor: 'pointer'
            }}
            title="Ajouter un logo d'entreprise pour qu'il apparaisse sur l'en-tête du document officiel"
          >
            🏢 {companyLogo ? "Logo Entreprise Défini" : "Ajouter Logo Entreprise"}
          </button>
          {companyLogo && (
            <button
              onClick={handleRemoveLogo}
              style={{
                background: 'rgba(239, 68, 68, 0.15)',
                border: '1px solid rgba(239, 68, 68, 0.3)',
                color: '#ef4444',
                padding: '6px 8px',
                borderRadius: '6px',
                fontSize: '0.75rem',
                cursor: 'pointer'
              }}
              title="Supprimer le logo actuel"
            >
              ✕ Retirer
            </button>
          )}

          <button
            onClick={handleSave}
            style={{
              background: 'rgba(56, 189, 248, 0.15)',
              border: '1px solid #38bdf8',
              color: '#38bdf8',
              padding: '6px 14px',
              borderRadius: '6px',
              fontSize: '0.8rem',
              fontWeight: 600,
              cursor: 'pointer'
            }}
          >
            💾 Sauvegarder
          </button>

          <button
            onClick={handleReset}
            style={{
              background: 'rgba(255, 255, 255, 0.05)',
              border: '1px solid rgba(255, 255, 255, 0.15)',
              color: '#94a3b8',
              padding: '6px 10px',
              borderRadius: '6px',
              fontSize: '0.8rem',
              cursor: 'pointer'
            }}
            title="Réinitialiser avec le contenu par défaut"
          >
            🔄 Réinit.
          </button>

          <button
            onClick={handlePrint}
            style={{
              background: 'linear-gradient(135deg, #0284c7, #2563eb)',
              border: 'none',
              color: '#fff',
              padding: '7px 16px',
              borderRadius: '6px',
              fontSize: '0.85rem',
              fontWeight: 700,
              display: 'flex',
              alignItems: 'center',
              gap: '6px',
              cursor: 'pointer',
              boxShadow: '0 4px 14px rgba(2, 132, 199, 0.4)'
            }}
          >
            📥 Exporter en PDF / Imprimer
          </button>
        </div>
      </div>

      {savedNotification && (
        <div className="no-print" style={{
          background: '#10b981',
          color: '#ffffff',
          padding: '8px 16px',
          textAlign: 'center',
          fontSize: '0.85rem',
          fontWeight: 600
        }}>
          ✓ {savedNotification}
        </div>
      )}

      {/* Main Workspace Body */}
      <div style={{ flex: 1, display: 'flex', overflow: 'hidden' }}>
        {/* Left / Input Form View */}
        {(viewMode === 'split' || viewMode === 'form') && (
          <div className="no-print" style={{
            flex: viewMode === 'split' ? '0 0 45%' : '1 1 100%',
            overflowY: 'auto',
            padding: '1.5rem',
            borderRight: viewMode === 'split' ? '1px solid rgba(255, 255, 255, 0.1)' : 'none',
            background: 'rgba(15, 23, 42, 0.8)'
          }}>
            <div style={{ maxWidth: '800px', margin: '0 auto' }}>
              <div style={{ marginBottom: '1.5rem' }}>
                <h3 style={{ margin: 0, fontSize: '1.15rem', color: '#38bdf8', display: 'flex', alignItems: 'center', gap: '8px' }}>
                  <span>✍️</span> Formulaire de Saisie du Document
                </h3>
                <p style={{ margin: '4px 0 0', color: '#94a3b8', fontSize: '0.82rem' }}>
                  Modifiez les données ci-dessous. L'aperçu officiel se met à jour en temps réel.
                </p>
              </div>

              {/* Logo preview card in form */}
              <div style={{
                background: 'rgba(30, 41, 59, 0.6)',
                border: '1px dashed rgba(255, 255, 255, 0.15)',
                borderRadius: '8px',
                padding: '1rem',
                marginBottom: '1.2rem',
                display: 'flex',
                alignItems: 'center',
                justifyContent: 'space-between'
              }}>
                <div style={{ display: 'flex', alignItems: 'center', gap: '1rem' }}>
                  {companyLogo ? (
                    <img src={companyLogo} alt="Logo Entreprise" style={{ maxHeight: '48px', maxWidth: '120px', objectFit: 'contain', background: '#fff', padding: '4px', borderRadius: '4px' }} />
                  ) : (
                    <div style={{ width: '48px', height: '48px', background: 'rgba(255,255,255,0.05)', borderRadius: '6px', display: 'flex', alignItems: 'center', justifyContent: 'center', color: '#64748b', fontSize: '1.4rem' }}>
                      🏢
                    </div>
                  )}
                  <div>
                    <div style={{ fontSize: '0.85rem', fontWeight: 700, color: '#f8fafc' }}>
                      {companyLogo ? "Logo de l'Entreprise Configuré" : "Aucun Logo Entreprise Défini"}
                    </div>
                    <div style={{ fontSize: '0.75rem', color: '#94a3b8' }}>
                      {companyLogo ? "Ce logo figurera automatiquement en tête de page de tous vos exports PDF." : "Ajoutez le logo de votre société pour personnaliser le document officiel."}
                    </div>
                  </div>
                </div>
                <button
                  onClick={() => fileInputRef.current?.click()}
                  style={{
                    background: '#0284c7',
                    border: 'none',
                    color: '#fff',
                    padding: '6px 12px',
                    borderRadius: '6px',
                    fontSize: '0.78rem',
                    fontWeight: 600,
                    cursor: 'pointer'
                  }}
                >
                  {companyLogo ? "Changer le logo" : "Parcourir..."}
                </button>
              </div>

              {/* Document Metadata Form */}
              <div style={{ background: 'rgba(30, 41, 59, 0.5)', borderRadius: '8px', padding: '1.2rem', marginBottom: '1.2rem', border: '1px solid rgba(255, 255, 255, 0.08)' }}>
                <div style={{ fontSize: '0.85rem', fontWeight: 700, color: '#e2e8f0', marginBottom: '0.8rem', textTransform: 'uppercase' }}>
                  Métadonnées du Document
                </div>
                <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '0.8rem', marginBottom: '0.8rem' }}>
                  <div>
                    <label style={{ display: 'block', fontSize: '0.75rem', color: '#94a3b8', marginBottom: '3px' }}>Titre du Document</label>
                    <input
                      type="text"
                      value={reportData.title}
                      onChange={(e) => setReportData({ ...reportData, title: e.target.value })}
                      style={{ width: '100%', background: 'rgba(15, 23, 42, 0.8)', border: '1px solid rgba(255,255,255,0.15)', color: '#fff', padding: '6px 10px', borderRadius: '4px', fontSize: '0.85rem' }}
                    />
                  </div>
                  <div>
                    <label style={{ display: 'block', fontSize: '0.75rem', color: '#94a3b8', marginBottom: '3px' }}>Sous-Titre / Objet</label>
                    <input
                      type="text"
                      value={reportData.subtitle}
                      onChange={(e) => setReportData({ ...reportData, subtitle: e.target.value })}
                      style={{ width: '100%', background: 'rgba(15, 23, 42, 0.8)', border: '1px solid rgba(255,255,255,0.15)', color: '#fff', padding: '6px 10px', borderRadius: '4px', fontSize: '0.85rem' }}
                    />
                  </div>
                </div>

                <div style={{ display: 'grid', gridTemplateColumns: 'repeat(4, 1fr)', gap: '0.8rem' }}>
                  <div>
                    <label style={{ display: 'block', fontSize: '0.75rem', color: '#94a3b8', marginBottom: '3px' }}>Réf. Code</label>
                    <input
                      type="text"
                      value={reportData.code}
                      onChange={(e) => setReportData({ ...reportData, code: e.target.value })}
                      style={{ width: '100%', background: 'rgba(15, 23, 42, 0.8)', border: '1px solid rgba(255,255,255,0.15)', color: '#fff', padding: '6px 10px', borderRadius: '4px', fontSize: '0.82rem' }}
                    />
                  </div>
                  <div>
                    <label style={{ display: 'block', fontSize: '0.75rem', color: '#94a3b8', marginBottom: '3px' }}>Version</label>
                    <input
                      type="text"
                      value={reportData.version}
                      onChange={(e) => setReportData({ ...reportData, version: e.target.value })}
                      style={{ width: '100%', background: 'rgba(15, 23, 42, 0.8)', border: '1px solid rgba(255,255,255,0.15)', color: '#fff', padding: '6px 10px', borderRadius: '4px', fontSize: '0.82rem' }}
                    />
                  </div>
                  <div>
                    <label style={{ display: 'block', fontSize: '0.75rem', color: '#94a3b8', marginBottom: '3px' }}>Date</label>
                    <input
                      type="date"
                      value={reportData.date}
                      onChange={(e) => setReportData({ ...reportData, date: e.target.value })}
                      style={{ width: '100%', background: 'rgba(15, 23, 42, 0.8)', border: '1px solid rgba(255,255,255,0.15)', color: '#fff', padding: '6px 10px', borderRadius: '4px', fontSize: '0.82rem' }}
                    />
                  </div>
                  <div>
                    <label style={{ display: 'block', fontSize: '0.75rem', color: '#94a3b8', marginBottom: '3px' }}>Statut</label>
                    <select
                      value={reportData.status}
                      onChange={(e) => setReportData({ ...reportData, status: e.target.value as any })}
                      style={{ width: '100%', background: 'rgba(15, 23, 42, 0.8)', border: '1px solid rgba(255,255,255,0.15)', color: '#fff', padding: '6px 10px', borderRadius: '4px', fontSize: '0.82rem' }}
                    >
                      <option value="Brouillon">Brouillon</option>
                      <option value="En revue">En revue</option>
                      <option value="Validé">Validé</option>
                      <option value="Diffusé">Diffusé</option>
                    </select>
                  </div>
                </div>

                <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '0.8rem', marginTop: '0.8rem' }}>
                  <div>
                    <label style={{ display: 'block', fontSize: '0.75rem', color: '#94a3b8', marginBottom: '3px' }}>Auteur / Rédacteur</label>
                    <input
                      type="text"
                      value={reportData.author}
                      onChange={(e) => setReportData({ ...reportData, author: e.target.value })}
                      style={{ width: '100%', background: 'rgba(15, 23, 42, 0.8)', border: '1px solid rgba(255,255,255,0.15)', color: '#fff', padding: '6px 10px', borderRadius: '4px', fontSize: '0.85rem' }}
                    />
                  </div>
                  <div>
                    <label style={{ display: 'block', fontSize: '0.75rem', color: '#94a3b8', marginBottom: '3px' }}>Instance / Approbateur</label>
                    <input
                      type="text"
                      value={reportData.approver}
                      onChange={(e) => setReportData({ ...reportData, approver: e.target.value })}
                      style={{ width: '100%', background: 'rgba(15, 23, 42, 0.8)', border: '1px solid rgba(255,255,255,0.15)', color: '#fff', padding: '6px 10px', borderRadius: '4px', fontSize: '0.85rem' }}
                    />
                  </div>
                </div>

                <div style={{ marginTop: '0.8rem' }}>
                  <label style={{ display: 'block', fontSize: '0.75rem', color: '#94a3b8', marginBottom: '3px' }}>Résumé Exécutif (Introduction)</label>
                  <textarea
                    rows={3}
                    value={reportData.summary}
                    onChange={(e) => setReportData({ ...reportData, summary: e.target.value })}
                    style={{ width: '100%', background: 'rgba(15, 23, 42, 0.8)', border: '1px solid rgba(255,255,255,0.15)', color: '#fff', padding: '8px 10px', borderRadius: '4px', fontSize: '0.85rem', resize: 'vertical' }}
                  />
                </div>
              </div>

              {/* Sections Editor */}
              <div style={{ display: 'flex', flexDirection: 'column', gap: '1.2rem' }}>
                {reportData.sections.map((section, sIdx) => (
                  <div key={section.id} style={{ background: 'rgba(30, 41, 59, 0.5)', borderRadius: '8px', padding: '1.2rem', border: '1px solid rgba(255, 255, 255, 0.08)' }}>
                    <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginBottom: '0.6rem' }}>
                      <input
                        type="text"
                        value={section.title}
                        onChange={(e) => {
                          const updated = { ...reportData };
                          updated.sections[sIdx].title = e.target.value;
                          setReportData(updated);
                        }}
                        style={{
                          background: 'transparent',
                          border: 'none',
                          borderBottom: '1px dashed rgba(255,255,255,0.2)',
                          color: '#38bdf8',
                          fontWeight: 700,
                          fontSize: '0.95rem',
                          width: '70%',
                          padding: '4px 0'
                        }}
                      />
                      <span style={{ fontSize: '0.72rem', color: '#94a3b8', background: 'rgba(255,255,255,0.06)', padding: '2px 8px', borderRadius: '4px' }}>
                        {section.type === 'table' ? 'Tableau' : 'Texte narratif'}
                      </span>
                    </div>

                    {/* Section Text */}
                    <textarea
                      rows={section.type === 'table' ? 2 : 4}
                      value={section.content}
                      onChange={(e) => updateSectionContent(sIdx, e.target.value)}
                      style={{
                        width: '100%',
                        background: 'rgba(15, 23, 42, 0.8)',
                        border: '1px solid rgba(255,255,255,0.15)',
                        color: '#f8fafc',
                        padding: '8px 10px',
                        borderRadius: '4px',
                        fontSize: '0.85rem',
                        resize: 'vertical',
                        marginBottom: section.type === 'table' ? '0.8rem' : '0'
                      }}
                    />

                    {/* Section Table if any */}
                    {section.type === 'table' && section.tableHeaders && section.tableRows && (
                      <div style={{ overflowX: 'auto', marginTop: '0.6rem' }}>
                        <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '0.8rem' }}>
                          <thead>
                            <tr style={{ background: 'rgba(15, 23, 42, 0.9)', color: '#94a3b8' }}>
                              {section.tableHeaders.map((th, hIdx) => (
                                <th key={hIdx} style={{ padding: '6px 8px', border: '1px solid rgba(255,255,255,0.1)', textAlign: 'left', fontWeight: 600 }}>
                                  {th}
                                </th>
                              ))}
                              <th style={{ width: '40px', padding: '6px 8px', border: '1px solid rgba(255,255,255,0.1)', textAlign: 'center' }}>
                                Act.
                              </th>
                            </tr>
                          </thead>
                          <tbody>
                            {section.tableRows.map((row, rIdx) => (
                              <tr key={rIdx} style={{ background: rIdx % 2 === 0 ? 'rgba(255,255,255,0.02)' : 'transparent' }}>
                                {row.map((cell, cIdx) => (
                                  <td key={cIdx} style={{ padding: '4px', border: '1px solid rgba(255,255,255,0.08)' }}>
                                    <input
                                      type="text"
                                      value={cell}
                                      onChange={(e) => updateTableCell(sIdx, rIdx, cIdx, e.target.value)}
                                      style={{
                                        width: '100%',
                                        background: 'transparent',
                                        border: 'none',
                                        color: '#fff',
                                        padding: '4px',
                                        fontSize: '0.78rem'
                                      }}
                                    />
                                  </td>
                                ))}
                                <td style={{ textAlign: 'center', border: '1px solid rgba(255,255,255,0.08)' }}>
                                  <button
                                    onClick={() => deleteTableRow(sIdx, rIdx)}
                                    style={{ background: 'none', border: 'none', color: '#ef4444', cursor: 'pointer', fontSize: '0.9rem' }}
                                    title="Supprimer la ligne"
                                  >
                                    ✕
                                  </button>
                                </td>
                              </tr>
                            ))}
                          </tbody>
                        </table>
                        <div style={{ marginTop: '6px', textAlign: 'right' }}>
                          <button
                            onClick={() => addTableRow(sIdx)}
                            style={{
                              background: 'rgba(56, 189, 248, 0.1)',
                              border: '1px solid rgba(56, 189, 248, 0.3)',
                              color: '#38bdf8',
                              padding: '4px 10px',
                              borderRadius: '4px',
                              fontSize: '0.75rem',
                              fontWeight: 600,
                              cursor: 'pointer'
                            }}
                          >
                            + Ajouter une ligne
                          </button>
                        </div>
                      </div>
                    )}
                  </div>
                ))}
              </div>

              {/* Signatures Editor */}
              <div style={{ background: 'rgba(30, 41, 59, 0.5)', borderRadius: '8px', padding: '1.2rem', marginTop: '1.2rem', border: '1px solid rgba(255, 255, 255, 0.08)' }}>
                <div style={{ fontSize: '0.85rem', fontWeight: 700, color: '#e2e8f0', marginBottom: '0.8rem', textTransform: 'uppercase' }}>
                  Signatures & Validations
                </div>
                <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(220px, 1fr))', gap: '0.8rem' }}>
                  {reportData.signatures.map((sig, sIdx) => (
                    <div key={sIdx} style={{ background: 'rgba(15, 23, 42, 0.6)', padding: '0.8rem', borderRadius: '6px', border: '1px solid rgba(255,255,255,0.06)' }}>
                      <input
                        type="text"
                        value={sig.role}
                        onChange={(e) => {
                          const updated = { ...reportData };
                          updated.signatures[sIdx].role = e.target.value;
                          setReportData(updated);
                        }}
                        style={{ width: '100%', background: 'transparent', border: 'none', color: '#38bdf8', fontWeight: 700, fontSize: '0.82rem', marginBottom: '4px' }}
                      />
                      <input
                        type="text"
                        value={sig.name}
                        onChange={(e) => {
                          const updated = { ...reportData };
                          updated.signatures[sIdx].name = e.target.value;
                          setReportData(updated);
                        }}
                        placeholder="Nom du signataire"
                        style={{ width: '100%', background: 'rgba(255,255,255,0.05)', border: '1px solid rgba(255,255,255,0.1)', color: '#fff', padding: '4px 6px', borderRadius: '4px', fontSize: '0.8rem', marginBottom: '4px' }}
                      />
                      <div style={{ display: 'flex', gap: '4px' }}>
                        <input
                          type="date"
                          value={sig.date}
                          onChange={(e) => {
                            const updated = { ...reportData };
                            updated.signatures[sIdx].date = e.target.value;
                            setReportData(updated);
                          }}
                          style={{ width: '60%', background: 'rgba(255,255,255,0.05)', border: '1px solid rgba(255,255,255,0.1)', color: '#fff', padding: '3px 4px', borderRadius: '4px', fontSize: '0.75rem' }}
                        />
                        <select
                          value={sig.status}
                          onChange={(e) => {
                            const updated = { ...reportData };
                            updated.signatures[sIdx].status = e.target.value as any;
                            setReportData(updated);
                          }}
                          style={{ width: '40%', background: 'rgba(255,255,255,0.05)', border: '1px solid rgba(255,255,255,0.1)', color: '#fff', padding: '3px 4px', borderRadius: '4px', fontSize: '0.72rem' }}
                        >
                          <option value="Approuvé">Approuvé</option>
                          <option value="Validé sous réserve">Sous réserve</option>
                          <option value="En attente">En attente</option>
                        </select>
                      </div>
                    </div>
                  ))}
                </div>
              </div>
            </div>
          </div>
        )}

        {/* Right / Official Document Preview & Print Sheet */}
        {(viewMode === 'split' || viewMode === 'preview') && (
          <div className="report-preview-scroll-area" style={{
            flex: viewMode === 'split' ? '0 0 55%' : '1 1 100%',
            overflowY: 'auto',
            padding: '2rem',
            background: '#334155',
            display: 'flex',
            justifyContent: 'center'
          }}>
            {/* The A4 Printable Paper */}
            <div className="printable-document" style={{
              width: '100%',
              maxWidth: '820px',
              minHeight: '1100px',
              background: '#ffffff',
              color: '#0f172a',
              boxShadow: '0 10px 40px rgba(0,0,0,0.4)',
              borderRadius: '4px',
              padding: '45px 50px',
              display: 'flex',
              flexDirection: 'column',
              fontFamily: '"Inter", -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif',
              boxSizing: 'border-box',
              position: 'relative'
            }}>
              {/* Document Header (Logo, Reference, Project Name) */}
              <div style={{
                display: 'flex',
                alignItems: 'center',
                justifyContent: 'space-between',
                borderBottom: '3px solid #0284c7',
                paddingBottom: '18px',
                marginBottom: '25px'
              }}>
                {/* Left: Enterprise Logo or Default App Stamp */}
                <div style={{ display: 'flex', alignItems: 'center', gap: '14px' }}>
                  {companyLogo ? (
                    <img
                      src={companyLogo}
                      alt="Logo de l'entreprise"
                      style={{ maxHeight: '65px', maxWidth: '180px', objectFit: 'contain' }}
                    />
                  ) : (
                    <div style={{ display: 'flex', alignItems: 'center', gap: '10px' }}>
                      <img src="/logo.png" alt="FAAB PROJET" style={{ width: '48px', height: '48px', borderRadius: '8px' }} />
                      <div>
                        <div style={{ fontWeight: 900, fontSize: '1.2rem', color: '#0f172a', letterSpacing: '-0.5px' }}>
                          FAAB PROJET
                        </div>
                        <div style={{ fontSize: '0.65rem', color: '#64748b', textTransform: 'uppercase', letterSpacing: '1px', fontWeight: 700 }}>
                          KIT CHEF PROJET
                        </div>
                      </div>
                    </div>
                  )}
                </div>

                {/* Right: Reference & Version Box */}
                <div style={{ textAlign: 'right' }}>
                  <div style={{
                    display: 'inline-block',
                    background: '#f1f5f9',
                    border: '1px solid #cbd5e1',
                    borderRadius: '4px',
                    padding: '4px 10px',
                    fontSize: '0.75rem',
                    fontWeight: 700,
                    color: '#0369a1',
                    marginBottom: '4px'
                  }}>
                    {reportData.code} • Rev. {reportData.version}
                  </div>
                  <div style={{ fontSize: '0.75rem', color: '#64748b' }}>
                    Date : <strong>{new Date(reportData.date).toLocaleDateString('fr-FR', { day: '2-digit', month: 'long', year: 'numeric' })}</strong>
                  </div>
                  <div style={{ fontSize: '0.72rem', color: '#0284c7', fontWeight: 600, textTransform: 'uppercase' }}>
                    Statut : {reportData.status}
                  </div>
                </div>
              </div>

              {/* Title & Category Banner */}
              <div style={{ marginBottom: '24px' }}>
                <div style={{
                  display: 'inline-block',
                  background: '#e0f2fe',
                  color: '#0369a1',
                  padding: '3px 10px',
                  borderRadius: '12px',
                  fontSize: '0.75rem',
                  fontWeight: 700,
                  marginBottom: '8px',
                  textTransform: 'uppercase',
                  letterSpacing: '0.5px'
                }}>
                  {reportData.category}
                </div>
                <h1 style={{ margin: '0 0 6px 0', fontSize: '1.75rem', fontWeight: 800, color: '#0f172a', lineHeight: 1.25 }}>
                  {reportData.title}
                </h1>
                <p style={{ margin: 0, fontSize: '0.95rem', color: '#475569', fontStyle: 'italic' }}>
                  {reportData.subtitle}
                </p>
              </div>

              {/* Document Identity Meta Table */}
              <div style={{
                background: '#f8fafc',
                border: '1px solid #e2e8f0',
                borderRadius: '6px',
                padding: '12px 16px',
                marginBottom: '24px',
                display: 'grid',
                gridTemplateColumns: 'repeat(3, 1fr)',
                gap: '12px',
                fontSize: '0.8rem'
              }}>
                <div>
                  <span style={{ color: '#64748b', display: 'block', fontSize: '0.7rem', textTransform: 'uppercase', fontWeight: 600 }}>Rédacteur / Auteur</span>
                  <strong style={{ color: '#0f172a' }}>{reportData.author}</strong>
                </div>
                <div>
                  <span style={{ color: '#64748b', display: 'block', fontSize: '0.7rem', textTransform: 'uppercase', fontWeight: 600 }}>Instance de Validation</span>
                  <strong style={{ color: '#0f172a' }}>{reportData.approver}</strong>
                </div>
                <div>
                  <span style={{ color: '#64748b', display: 'block', fontSize: '0.7rem', textTransform: 'uppercase', fontWeight: 600 }}>Périmètre & Cadre</span>
                  <strong style={{ color: '#0f172a' }}>Kit Chef Projet • Document Officiel</strong>
                </div>
              </div>

              {/* Executive Summary */}
              {reportData.summary && (
                <div style={{
                  background: '#f0fdf4',
                  borderLeft: '4px solid #16a34a',
                  padding: '14px 18px',
                  borderRadius: '0 6px 6px 0',
                  marginBottom: '26px'
                }}>
                  <div style={{ fontSize: '0.75rem', fontWeight: 800, color: '#166534', textTransform: 'uppercase', letterSpacing: '0.5px', marginBottom: '4px' }}>
                    Synthèse & Objet du Document
                  </div>
                  <div style={{ fontSize: '0.88rem', color: '#14532d', lineHeight: 1.5, whiteSpace: 'pre-line' }}>
                    {reportData.summary}
                  </div>
                </div>
              )}

              {/* Dynamic Content Sections */}
              <div style={{ display: 'flex', flexDirection: 'column', gap: '22px', flex: 1 }}>
                {reportData.sections.map((section) => (
                  <div key={section.id} className="report-doc-section" style={{ pageBreakInside: 'avoid' }}>
                    <h2 style={{
                      margin: '0 0 10px 0',
                      fontSize: '1.15rem',
                      fontWeight: 700,
                      color: '#0f172a',
                      borderBottom: '1px solid #e2e8f0',
                      paddingBottom: '5px'
                    }}>
                      {section.title}
                    </h2>

                    {section.content && (
                      <div style={{
                        fontSize: '0.88rem',
                        lineHeight: 1.6,
                        color: '#334155',
                        marginBottom: section.type === 'table' ? '12px' : '0',
                        whiteSpace: 'pre-line'
                      }}>
                        {section.content}
                      </div>
                    )}

                    {/* Table Render */}
                    {section.type === 'table' && section.tableHeaders && section.tableRows && (
                      <table style={{
                        width: '100%',
                        borderCollapse: 'collapse',
                        fontSize: '0.82rem',
                        marginTop: '8px',
                        border: '1px solid #cbd5e1'
                      }}>
                        <thead>
                          <tr style={{ background: '#f1f5f9', color: '#0f172a' }}>
                            {section.tableHeaders.map((th, hIdx) => (
                              <th key={hIdx} style={{
                                padding: '8px 10px',
                                border: '1px solid #cbd5e1',
                                textAlign: 'left',
                                fontWeight: 700
                              }}>
                                {th}
                              </th>
                            ))}
                          </tr>
                        </thead>
                        <tbody>
                          {section.tableRows.map((row, rIdx) => (
                            <tr key={rIdx} style={{ background: rIdx % 2 === 1 ? '#f8fafc' : '#ffffff' }}>
                              {row.map((cell, cIdx) => (
                                <td key={cIdx} style={{
                                  padding: '8px 10px',
                                  border: '1px solid #e2e8f0',
                                  color: '#334155'
                                }}>
                                  {cell}
                                </td>
                              ))}
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    )}
                  </div>
                ))}
              </div>

              {/* Signatures & Formal Approvals Section */}
              <div style={{ marginTop: '35px', paddingTop: '20px', borderTop: '2px dashed #cbd5e1', pageBreakInside: 'avoid' }}>
                <div style={{ fontSize: '0.85rem', fontWeight: 800, color: '#0f172a', textTransform: 'uppercase', marginBottom: '14px', letterSpacing: '0.5px' }}>
                  Visa, Approbations & Signatures Contractuelles
                </div>

                <div style={{
                  display: 'grid',
                  gridTemplateColumns: `repeat(${Math.min(reportData.signatures.length, 3)}, 1fr)`,
                  gap: '16px'
                }}>
                  {reportData.signatures.map((sig, sIdx) => (
                    <div key={sIdx} style={{
                      border: '1px solid #cbd5e1',
                      borderRadius: '6px',
                      padding: '12px',
                      background: '#f8fafc',
                      display: 'flex',
                      flexDirection: 'column',
                      minHeight: '110px',
                      justifyContent: 'space-between'
                    }}>
                      <div>
                        <div style={{ fontSize: '0.72rem', color: '#64748b', textTransform: 'uppercase', fontWeight: 700 }}>
                          {sig.role}
                        </div>
                        <div style={{ fontSize: '0.9rem', fontWeight: 700, color: '#0f172a', marginTop: '2px' }}>
                          {sig.name}
                        </div>
                        <div style={{ fontSize: '0.72rem', color: '#0284c7', marginTop: '2px' }}>
                          {sig.status} • {new Date(sig.date).toLocaleDateString('fr-FR')}
                        </div>
                      </div>
                      <div style={{
                        marginTop: '15px',
                        paddingTop: '8px',
                        borderTop: '1px solid #e2e8f0',
                        fontSize: '0.68rem',
                        color: '#94a3b8',
                        fontFamily: 'monospace'
                      }}>
                        [Signature électronique / Visa]
                      </div>
                    </div>
                  ))}
                </div>
              </div>

              {/* Document Official Footer */}
              <div style={{
                marginTop: '35px',
                paddingTop: '14px',
                borderTop: '1px solid #e2e8f0',
                display: 'flex',
                alignItems: 'center',
                justifyContent: 'space-between',
                fontSize: '0.72rem',
                color: '#94a3b8'
              }}>
                <div>
                  Document certifié conforme • Réf: {reportData.code} • Généré le {new Date().toLocaleDateString('fr-FR')}
                </div>
                <div>
                  Page 1 / 1
                </div>
              </div>
            </div>
          </div>
        )}
      </div>
    </div>
  );
};
