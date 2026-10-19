import React, { useState } from 'react';

export interface TrainingSlide {
  slideNumber: number;
  title: string;
  category: string;
  badge?: string;
  subtitle?: string;
  bullets: string[];
  keyTakeaway?: string;
  structuredData?: { label: string; value: string; color?: string }[];
  diagramSteps?: { step: number; title: string; desc: string; icon: string }[];
}

export interface TrainingCourse {
  id: string;
  name: string;
  code: string;
  totalSlides: number;
  durationMinutes: number;
  level: 'Fondamentaux' | 'Intermédiaire' | 'Avancé' | 'Gouvernance';
  description: string;
  slides: TrainingSlide[];
}

export const trainingCourses: Record<string, TrainingCourse> = {
  'ppt-lancement': {
    id: 'ppt-lancement',
    name: 'Réunion de Lancement de Projet (Kick-off)',
    code: 'FORM-KICKOFF-08',
    totalSlides: 12,
    durationMinutes: 45,
    level: 'Fondamentaux',
    description: "Support de formation complet pour préparer, animer et réussir le kick-off d'un projet : alignement des équipes, partage des objectifs, gouvernance, jalons et règles de jeu collectives.",
    slides: [
      {
        slideNumber: 1,
        title: 'Réunion de Lancement de Projet (Kick-off Meeting)',
        category: 'Introduction & Cadrage',
        badge: 'Jalon J0',
        subtitle: 'Donner le coup d\'envoi officiel et fédérer toutes les parties prenantes',
        bullets: [
          'Le Kick-off matérialise l\'entrée en phase opérationnelle du projet.',
          'Objectif clé : s\'assurer que l\'ensemble des acteurs partagent la même vision, les mêmes priorités et connaissent précisément leur rôle.',
          'Présentation formelle du sponsor, du chef de projet et de l\'équipe cœur.',
          'Instauration d\'un climat de confiance et d\'engagement collaboratif dès le premier jour.'
        ],
        keyTakeaway: "Un kick-off réussi réduit de 40% les incompréhensions de périmètre et renforce l'adhésion dès les premières semaines.",
        structuredData: [
          { label: 'Durée idéale', value: '1h30 à 2h00', color: '#38bdf8' },
          { label: 'Public cible', value: 'Équipe projet + Sponsor + Métiers', color: '#10b981' },
          { label: 'Livrable en sortie', value: 'CR Kick-off & Charte signée', color: '#f59e0b' }
        ]
      },
      {
        slideNumber: 2,
        title: 'Ordre du Jour & Contexte Stratégique',
        category: 'Alignement Stratégique',
        badge: 'Vision Globale',
        subtitle: 'Pourquoi lançons-nous ce projet maintenant ?',
        bullets: [
          'Contexte d\'entreprise : opportunités de marché, modernisation des processus ou mise en conformité réglementaire.',
          'Problématique adressée : les limites et coûts de fonctionnement actuels.',
          'Bénéfices cibles : gains d\'efficacité, diminution du taux d\'erreurs, meilleure satisfaction client.',
          'Alignement direct avec les 3 priorités stratégiques annuelles de l\'organisation.'
        ],
        keyTakeaway: "Chaque membre de l'équipe doit pouvoir résumer en une phrase l'intérêt stratégique du projet pour l'entreprise."
      },
      {
        slideNumber: 3,
        title: 'Périmètre du Projet : In Scope vs Out of Scope',
        category: 'Maîtrise du Périmètre',
        badge: 'Frontières Claires',
        subtitle: 'Fixer des limites imperméables pour éviter la dérive de périmètre (Scope Creep)',
        bullets: [
          'IN SCOPE (Ce qui est formellement engagé) : Socle applicatif, intégration des modèles de rapports, export PDF vectoriel, visualiseurs de formation.',
          'OUT OF SCOPE (Ce qui est formellement exclu ou reporté) : Développement de passerelles propriétaires avec anciens ERP obsolètes.',
          'Toute demande complémentaire doit obligatoirement suivre le processus formel de demande de changement (Workflow de modification).'
        ],
        keyTakeaway: "Définir ce qu'on ne fera PAS est tout aussi crucial que définir ce qu'on fera.",
        structuredData: [
          { label: 'Inclus (In-Scope)', value: '100% des attendus validés', color: '#10b981' },
          { label: 'Exclus (Out-Scope)', value: 'Garantit le respect du budget', color: '#ef4444' }
        ]
      },
      {
        slideNumber: 4,
        title: 'Gouvernance du Projet & Matrice RACI',
        category: 'Organisation & Rôles',
        badge: 'Clarification des Rôles',
        subtitle: 'Qui décide, qui fait, qui est consulté, qui est informé ?',
        bullets: [
          'R - Réalisateur (Responsible) : L\'équipe opérationnelle chargée d\'exécuter la tâche.',
          'A - Approbateur (Accountable) : Le décideur unique qui valide le livrable final (1 seul A par tâche !).',
          'C - Consulté (Consulted) : Les experts métiers ou techniques sollicités pour avis.',
          'I - Informé (Informed) : Les acteurs tenus au courant des avancements sans intervention directe.'
        ],
        keyTakeaway: "Une tâche sans 'A' est orpheline ; une tâche avec plusieurs 'A' génère des conflits d'arbitrage.",
        diagramSteps: [
          { step: 1, title: 'Sponsor', desc: 'Finance et arbitre en dernier ressort', icon: '👑' },
          { step: 2, title: 'Chef de Projet', desc: 'Pilote délais, coûts et qualité', icon: '🎯' },
          { step: 3, title: 'Équipe Métier', desc: 'Apporte l\'expertise fonctionnelle', icon: '💼' },
          { step: 4, title: 'Équipe Technique', desc: 'Réalise et qualifie la solution', icon: '⚙️' }
        ]
      },
      {
        slideNumber: 5,
        title: 'Macro-Planning & Jalons Clés Contractuels',
        category: 'Planification',
        badge: 'Trajectoire Temps',
        subtitle: 'Les rendez-vous décisifs tout au long du cycle de vie',
        bullets: [
          'Jalon J0 : Lancement officiel et validation de la charte de projet.',
          'Jalon J1 : Approbation du cahier des charges fonctionnel et de l\'architecture cible.',
          'Jalon J2 : Fin des développements et livraison pour recette.',
          'Jalon J3 : Signature sans réserve du Procès-Verbal de Recette Utilisateur (UAT).',
          'Jalon J4 : Déploiement généralisé et clôture définitive du projet.'
        ],
        keyTakeaway: "Un jalon n'est pas une simple date : c'est un livrable audité et approuvé formellement."
      },
      {
        slideNumber: 6,
        title: 'Budget Prévisionnel & Pilotage Financier',
        category: 'Gestion Financière',
        badge: 'Ressources & Coûts',
        subtitle: 'Ventilation analytique et réserve pour aléas',
        bullets: [
          'Enveloppe globale allouée : 85 000 € HT.',
          'Postes majeurs : Réalisation technique (45%), Conception et spécifications (18%), Tests & Recette (14%), Accompagnement (12%).',
          'Réserve pour risques & imprévus : 11% sanctuarisés sous l\'autorité du chef de projet.',
          'Suivi par la Valeur Acquise (Earned Value Management) pour mesurer la rentabilité du capital engagé.'
        ],
        keyTakeaway: "Ne jamais consommer la réserve d'aléas pour financer des demandes de changement non budgétées."
      },
      {
        slideNumber: 7,
        title: 'Cartographie des Risques Majeurs & Réponses',
        category: 'Management des Risques',
        badge: 'Anticipation',
        subtitle: 'Identifier les dangers potentiels avant qu\'ils ne se transforment en problèmes',
        bullets: [
          'Risque 1 - Disponibilité des experts métiers pendant la recette : Planifié 4 semaines à l\'avance dans les agendas.',
          'Risque 2 - Dérive des demandes fonctionnelles : Filtrage strict par le processus de changement.',
          'Risque 3 - Dépendance technique externe : Ateliers de convergence bi-hebdomadaires.',
          'Chaque risque est doté d\'une criticité (P × I) et d\'un plan d\'atténuation préventif.'
        ],
        keyTakeaway: "Un bon chef de projet ne résout pas les crises : il empêche les risques de devenir des crises."
      },
      {
        slideNumber: 8,
        title: 'Plan de Communication & Rituels d\'Équipe',
        category: 'Communication',
        badge: 'Synchronisation',
        subtitle: 'Les rendez-vous réguliers de coordination et de reporting',
        bullets: [
          'Comité de Projet (COPROJ) : Hebdomadaire (1h) - Revue des actions, blocages et tâches à 7 jours.',
          'Comité de Pilotage (COPIL) : Mensuel (1h30) - Décisions stratégiques, revue budgétaire et arbitrages.',
          'Rapport Flash Hebdo : Diffusé chaque vendredi à 17h avec météo projet.',
          'Outil unique de traçabilité : Plateforme de pilotage intégrée pour éviter la dispersion des fichiers.'
        ],
        keyTakeaway: "Trop d'information tue l'information : privilégier des formats synthétiques et standardisés."
      },
      {
        slideNumber: 9,
        title: 'Facteurs Clés de Succès (FCS) & Règles du Jeu',
        category: 'Excellence Opérationnelle',
        badge: 'Culture Projet',
        subtitle: 'Les comportements et pratiques qui garantissent la réussite',
        bullets: [
          'Transparence immédiate : signalez une difficulté dès qu\'elle survient plutôt qu\'à la dernière minute.',
          'Respect strict des engagements de livraison et des dates convenues.',
          'Documentation continue au fil de l\'eau (pas d\'effet tunnel).',
          'Bienveillance exigeante : focus sur les solutions et non sur la recherche de coupables.'
        ],
        keyTakeaway: "Dans un projet, les mauvaises nouvelles doivent voyager plus vite que les bonnes nouvelles."
      },
      {
        slideNumber: 10,
        title: 'Prochaines Actions Immédiates (Plan à 15 Jours)',
        category: 'Mise en Action',
        badge: 'Feuille de Route Court Terme',
        subtitle: 'Ce qui doit être fait d\'ici la prochaine séance',
        bullets: [
          'Diffusion du compte-rendu officiel du Kick-off à tous les participants sous 24h.',
          'Recueil des signatures de la Charte de Projet.',
          'Mise à disposition des accès applicatifs pour tous les intervenants.',
          'Démarrage des ateliers détaillés de recueil des exigences fonctionnelles.'
        ],
        keyTakeaway: "Un bon lancement se mesure à la vitesse d'exécution des décisions prises en kick-off."
      },
      {
        slideNumber: 11,
        title: 'Tour de Table & Questions / Réponses',
        category: 'Échanges Ouverts',
        badge: 'Partage',
        subtitle: 'Lever les doutes et recueillir les premières impressions',
        bullets: [
          'Expression libre de chaque intervenant sur ses attentes et contraintes.',
          'Validation collégiale du calendrier prévisionnel des réunions.',
          'Clarification des zones d\'ombre potentielles.',
          'Confirmation formelle de l\'adhésion de tous au projet.'
        ],
        keyTakeaway: "S'assurer que personne ne quitte la pièce avec une question non formulée ou un doute inexprimé."
      },
      {
        slideNumber: 12,
        title: 'Conclusion & Clôture du Kick-off',
        category: 'Conclusion',
        badge: 'En Marche !',
        subtitle: 'Le projet est officiellement sur les rails !',
        bullets: [
          'Mandat formel accordé au chef de projet.',
          'Équipe mobilisée et motivée vers un objectif commun clair.',
          'Rendez-vous fixé pour le premier point de projet COPROJ la semaine prochaine.',
          'FAAB PROJET : Vos outils et modèles sont d\'ores et déjà opérationnels pour vous accompagner.'
        ],
        keyTakeaway: "L'aventure commence. Restons rigoureux, agiles et focalisés sur la valeur livrée !"
      }
    ]
  },

  'ppt-pert': {
    id: 'ppt-pert',
    name: 'Méthodologie Réseau PERT & Chemin Critique',
    code: 'FORM-PERT-14',
    totalSlides: 7,
    durationMinutes: 30,
    level: 'Avancé',
    description: "Formation technique approfondie sur la méthode PERT (Program Evaluation and Review Technique) : calculs des dates au plus tôt, au plus tard, marges libres, marges totales et chemin critique.",
    slides: [
      {
        slideNumber: 1,
        title: 'Méthodologie Réseau PERT & Chemin Critique',
        category: 'Introduction PERT',
        badge: 'Théorie & Pratique',
        subtitle: 'Modéliser les dépendances pour maîtriser la durée incompressible du projet',
        bullets: [
          'Le réseau PERT organise logiquement les tâches sous forme de graphe orienté.',
          'Il permet de visualiser immédiatement les antécédences et les tâches pouvant être exécutées en parallèle.',
          'Il identifie de manière mathématique le Chemin Critique : la suite de tâches dont tout retard repousse la fin du projet.',
          'Complément indispensable du diagramme de Gantt pour les projets complexes.'
        ],
        keyTakeaway: "Le PERT répond à la question clé : 'Quelle est la durée minimale théorique absolue de mon projet ?'",
        structuredData: [
          { label: 'Création', value: 'Marine US (Projet Polaris)', color: '#38bdf8' },
          { label: 'Concept clé', value: 'Chemin Critique (Marge Nulle)', color: '#ef4444' },
          { label: 'Indicateur phare', value: 'Marge Totale & Marge Libre', color: '#10b981' }
        ]
      },
      {
        slideNumber: 2,
        title: 'Étape 1 : Liste des Tâches & Détermination des Antécédences',
        category: 'Méthode PERT',
        badge: 'Tableau des Antécédents',
        subtitle: 'Structurer rigoureusement les prérequis de chaque activité',
        bullets: [
          'Chaque tâche dispose d\'un code unique (A, B, C...) et d\'une durée estimée en jours ou semaines.',
          'On identifie pour chaque tâche ses tâches immédiatement antérieures (qui doivent être achevées avant de démarrer).',
          'Exemple : La tâche D (Recette) a pour antécédent obligatoire la tâche C (Développement).',
          'Les tâches sans antécédents démarrent au temps T = 0.'
        ],
        keyTakeaway: "Une erreur dans le tableau d'antécédences fausse l'intégralité du calcul du chemin critique."
      },
      {
        slideNumber: 3,
        title: 'Étape 2 : Calcul des Dates au Plus Tôt (Passage Avant)',
        category: 'Calcul PERT',
        badge: 'De Gauche à Droite',
        subtitle: 'Au plus tôt, quand une tâche peut-elle commencer et se terminer ?',
        bullets: [
          'On démarre du début du projet (Date début = 0).',
          'Date Fin au Plus Tôt = Date Début au Plus Tôt + Durée de la tâche.',
          'Règle d\'or pour convergence (quand plusieurs tâches précèdent une même étape) : on retient le MAXIMUM des dates de fin au plus tôt des tâches antérieures.',
          'La date de fin au plus tôt de la dernière tâche donne la durée totale du projet.'
        ],
        keyTakeaway: "Règle de convergence passage avant : TOUJOURS prendre la valeur la plus ÉLEVÉE (le goulot d'étranglement)."
      },
      {
        slideNumber: 4,
        title: 'Étape 3 : Calcul des Dates au Plus Tard (Passage Arrière)',
        category: 'Calcul PERT',
        badge: 'De Droite à Gauche',
        subtitle: 'Jusqu\'à quand une tâche peut-elle tarder sans retarder le projet final ?',
        bullets: [
          'On part de la fin du projet (Date fin au plus tard = Date fin au plus tôt du projet).',
          'Date Début au Plus Tard = Date Fin au Plus Tard - Durée de la tâche.',
          'Règle d\'or pour divergence arrière : lorsque plusieurs tâches partent d\'une même étape, on retient le MINIMUM des dates de début au plus tard.',
          'On remonte ainsi jusqu\'au début du projet (Date = 0).'
        ],
        keyTakeaway: "Règle passage arrière : TOUJOURS prendre la valeur la plus FAIBLE pour ne pas pénaliser les tâches suivantes."
      },
      {
        slideNumber: 5,
        title: 'Étape 4 : Calcul des Marges Totales & Marges Libres',
        category: 'Analyse des Marges',
        badge: 'Flexibilité',
        subtitle: 'Quantifier le degré de liberté dont dispose le chef de projet',
        bullets: [
          'Marge Totale (MT) = Date Début au Plus Tard - Date Début au Plus Tôt (ou Fin Tard - Fin Tôt). C\'est le retard qu\'une tâche peut accumuler sans repousser la fin du projet.',
          'Marge Libre (ML) = Date Début au Plus Tôt des tâches suivantes - Fin au Plus Tôt. C\'est le retard qu\'une tâche peut prendre sans impacter le démarrage de ses suivantes.',
          'Une tâche avec MT = 5 jours peut glisser de 5 jours sans danger pour la livraison finale.'
        ],
        keyTakeaway: "La Marge Libre est toujours inférieure ou égale à la Marge Totale (ML ≤ MT)."
      },
      {
        slideNumber: 6,
        title: 'Étape 5 : Détection du Chemin Critique (Marge Nulle)',
        category: 'Chemin Critique',
        badge: 'Alerte Rouge',
        subtitle: 'Les tâches sous surveillance absolue',
        bullets: [
          'Le Chemin Critique est constitué de l\'ensemble des tâches dont la Marge Totale est ÉGALE À ZÉRO (MT = 0).',
          'Tout retard d\'une seule journée sur une tâche critique repousse mathématiquement la fin du projet d\'une journée.',
          'Les tâches critiques n\'ont aucune flexibilité.',
          'C\'est sur ces tâches que le chef de projet doit concentrer ses meilleures ressources et sa vigilance quotidienne.'
        ],
        keyTakeaway: "Un projet peut avoir plusieurs chemins critiques simultanés si plusieurs branches ont la même durée maximale."
      },
      {
        slideNumber: 7,
        title: 'Pilotage & Compression du Planning (Crashing / Fast-tracking)',
        category: 'Optimisation Planning',
        badge: 'Stratégie',
        subtitle: 'Comment gagner du temps lorsque le calendrier est trop serré ?',
        bullets: [
          'Fast-tracking : Paralléliser des tâches qui étaient prévues en séquence (augmente les risques).',
          'Crashing : Allouer des ressources supplémentaires sur les tâches DU CHEMIN CRITIQUE uniquement (inutile d\'accélérer une tâche non critique !).',
          'Attention : Accélérer le chemin critique peut faire émerger un nouveau chemin critique secondaire.',
          'Mettre à jour le PERT à chaque revue d\'avancement pour recalculer les marges résiduelles.'
        ],
        keyTakeaway: "Pour réduire la durée d'un projet, il est INUTILE de réduire une tâche qui possède de la marge."
      }
    ]
  },

  'ppt-workflow-changement': {
    id: 'ppt-workflow-changement',
    name: 'Workflow des Demandes de Modification',
    code: 'FORM-CHANGE-25',
    totalSlides: 6,
    durationMinutes: 25,
    level: 'Intermédiaire',
    description: "Guide méthodologique sur la gestion des évolutions et demandes de changement : qualification, chiffrage d'impact, circuit d'approbation et mise à jour des baselines.",
    slides: [
      {
        slideNumber: 1,
        title: 'Gestion des Demandes de Changement (Change Management)',
        category: 'Principes Clés',
        badge: 'Gouvernance',
        subtitle: 'Protéger le triangle d\'or Coût - Délai - Qualité',
        bullets: [
          'Aucun projet ne se déroule à 100% comme prévu : le changement est naturel et inévitable.',
          'Le danger n\'est pas le changement en soi, mais le changement non maîtrisé, informel et non chiffré.',
          'Objectif du workflow : canaliser toute demande via un processus formel et traçable avant toute exécution.',
          'Garantit que le sponsor valide en connaissance de cause l\'impact financier et calendaire.'
        ],
        keyTakeaway: "Si ce n'est pas documenté dans une fiche de changement, ce n'est pas réalisable."
      },
      {
        slideNumber: 2,
        title: 'Étape 1 : Émission & Description du Besoin',
        category: 'Cycle de Vie',
        badge: 'Fiche RFC',
        subtitle: 'Formaliser par écrit l\'origine et la justification de la modification',
        bullets: [
          'Tout demandeur (client, métier, technique) remplit une Demande de Changement (Request For Change - RFC).',
          'Données requises : motif fonctionnel, valeur ajoutée attendue, degré d\'urgence.',
          'Le chef de projet enregistre la demande dans le Registre des Changements avec un identifiant unique (CHG-01).',
          'Vérification préalable : est-ce réellement une modification ou un simple ajustement compris dans le périmètre initial ?'
        ],
        keyTakeaway: "Exiger un motif métier : refuser les demandes formulées sous le simple prétexte du 'ce serait bien d'avoir'."
      },
      {
        slideNumber: 3,
        title: 'Étape 2 : Analyse d\'Impact Quadruple (Coût, Délai, Qualité, Risques)',
        category: 'Cycle de Vie',
        badge: 'Chiffrage d\'Impact',
        subtitle: 'Mesurer précisément les conséquences de l\'intégration de la demande',
        bullets: [
          'Impact Délais : Combien de jours de glissement sur les jalons clés ?',
          'Impact Budgétaire : Quel coût additionnel en jours/hommes et prestations ?',
          'Impact Qualité & Architecture : Quels sont les effets de bord sur les modules déjà livrés ?',
          'Impact Risques : La modification introduit-elle de nouveaux facteurs de vulnérabilité ?'
        ],
        keyTakeaway: "Ne jamais donner un accord de principe sans avoir posé par écrit les 4 composantes de l'impact."
      },
      {
        slideNumber: 4,
        title: 'Étape 3 : Circuit d\'Arbitrage & Décision du COPIL',
        category: 'Gouvernance',
        badge: 'Instances Décisionnelles',
        subtitle: 'Qui est habilité à accepter ou refuser une modification ?',
        bullets: [
          'Impact Mineur (< 2 j/h et 0 dérive jalon) : Pouvoir de validation délégué au Chef de Projet.',
          'Impact Majeur (> 2 j/h ou impact jalon / budget) : Décision obligatoire du Comité de Pilotage (COPIL).',
          '3 décisions possibles : Approuvée avec avenant, Ajournée pour étude approfondie, ou Rejetée avec motif consigné.',
          'La décision est notifiée par écrit au demandeur sous 48h.'
        ],
        keyTakeaway: "Le chef de projet éclaire la décision mais ne prend pas sur son dos les dépassements de budget décidés par les métiers."
      },
      {
        slideNumber: 5,
        title: 'Étape 4 : Révision des Baselines & Plan d\'Action',
        category: 'Mise en Œuvre',
        badge: 'Actualisation',
        subtitle: 'Intégrer formellement la modification approuvée dans le projet',
        bullets: [
          'Mise à jour du planning de référence (Baseline de planning modifiée).',
          'Ajustement du budget alloué avec signature d\'un avenant financier si nécessaire.',
          'Intégration des nouvelles tâches dans le plan de charge de l\'équipe.',
          'Communication à toutes les parties prenantes de la nouvelle date cible officielle.'
        ],
        keyTakeaway: "Une fois validée, la demande de changement devient partie intégrante du contrat de projet."
      },
      {
        slideNumber: 6,
        title: 'Bonnes Pratiques & Erreurs à Éviter',
        category: 'Synthèse',
        badge: 'Checklist Chef de Projet',
        subtitle: 'Les pièges classiques dans la gestion des évolutions',
        bullets: [
          'Piège 1 : Accepter des demandes orales lors d\'un couloir ou d\'un appel informel.',
          'Piège 2 : Réaliser le travail avant que la décision du COPIL ne soit actée et signée.',
          'Piège 3 : Oublier de mettre à jour le cahier de recette après une modification fonctionnelle.',
          'Règle d\'or : La fermeté bienveillante protège le projet, l\'équipe et le sponsor.'
        ],
        keyTakeaway: "Dire 'Oui mais voici ce que cela coûte et décale' est infiniment plus professionnel que dire 'Non' ou dire 'Oui' sans réfléchir."
      }
    ]
  },

  'ppt-rapport-flash': {
    id: 'ppt-rapport-flash',
    name: 'Reporting & Rapport Flash Projet',
    code: 'FORM-FLASH-30',
    totalSlides: 5,
    durationMinutes: 20,
    level: 'Fondamentaux',
    description: "Comment concevoir un rapport flash percutant pour le management : météo projet, jalons clés, indicateurs synthétiques, alertes et décisions attendues.",
    slides: [
      {
        slideNumber: 1,
        title: 'Le Rapport Flash : L\'Art du Reporting Synthétique',
        category: 'Principes',
        badge: '1 Page Unique',
        subtitle: 'Informer les décideurs en 2 minutes de lecture maximum',
        bullets: [
          'Les membres de la direction n\'ont pas le temps de lire des rapports de 30 pages.',
          'Le rapport flash est une synthèse visuelle sur une seule page, publiée à fréquence fixe (hebdomadaire ou bi-mensuelle).',
          'Il donne une vue instantanée sur la santé du projet, les réalisations récentes et les alertes critiques.',
          'C\'est le thermomètre de confiance entre le chef de projet et sa hiérarchie.'
        ],
        keyTakeaway: "Un décideur doit comprendre en 30 secondes chrono si le projet va bien ou s'il doit intervenir."
      },
      {
        slideNumber: 2,
        title: 'La Météo Projet à 4 Dimensions',
        category: 'Indicateurs Clés',
        badge: 'Feux Tricolores',
        subtitle: 'Évaluer objectivement la situation selon 4 piliers',
        bullets: [
          '☀️ Météo Délais : Respect du calendrier et des jalons contractuels.',
          '💰 Météo Coûts : Respect du budget consommé versus valeur livrée.',
          '🎯 Météo Périmètre / Qualité : Conformité des livrables et absence de régressions.',
          '🤝 Météo Climat / Équipe : Niveau d\'engagement, disponibilité et moral des intervenants.',
          'Code couleur standard : Vert (Nominal), Orange (Vigilance / Action engagée), Rouge (Crise / Décision requise).'
        ],
        keyTakeaway: "Ne jamais hésiter à passer un feu au Rouge : c'est un appel à l'aide managériale, pas un aveu d'échec."
      },
      {
        slideNumber: 3,
        title: 'Faits Marquants & Réalisations de la Période',
        category: 'Contenu',
        badge: 'Ce qui a été fait',
        subtitle: 'Valoriser les avancées concrètes et les victoires intermédiaires',
        bullets: [
          'Lister 3 à 5 réalisations tangibles majeures de la semaine écoulée.',
          'Mentionner les livrables validés et les recettes achevées.',
          'Éviter le jargon technique incompréhensible pour les profils métiers.',
          'Mettre en avant l\'atteinte des étapes significatives pour maintenir la dynamique.'
        ],
        keyTakeaway: "Décrivez des RÉSULTATS obtenus ('Module export PDF validé') plutôt que des ACTIVITÉS ('On a travaillé sur le code')."
      },
      {
        slideNumber: 4,
        title: 'Jalons à Venir & Prochaines Échéances Clés',
        category: 'Contenu',
        badge: 'Vision 30 Jours',
        subtitle: 'Donner de la visibilité sur les engagements court terme',
        bullets: [
          'Présenter les 3 prochains jalons décisionnels avec leur date prévisionnelle.',
          'Mettre en évidence les dates cibles des livrables en cours de finalisation.',
          'Préciser si la tendance est stable, en avance ou en risque de glissement.',
          'Permet aux décideurs d\'anticiper leurs créneaux de validation et revues.'
        ],
        keyTakeaway: "La visibilité à 30 jours désamorce le stress managérial et crédibilise le pilotage."
      },
      {
        slideNumber: 5,
        title: 'Points d\'Attention & Arbitrages Requis',
        category: 'Passage à l\'Action',
        badge: 'Décisions Immédiates',
        subtitle: 'Le levier d\'action managériale du chef de projet',
        bullets: [
          'Formuler clairement les points de blocage qui dépassent la compétence du chef de projet.',
          'Pour chaque point, proposer 2 options d\'arbitrage avec avantage / inconvénient.',
          'Indiquer la date limite à laquelle la décision doit être prise pour ne pas impacter le jalon.',
          'Transformer le rapport de simple outil d\'information en un instrument d\'action.'
        ],
        keyTakeaway: "Ne venez jamais avec un problème sans proposer au moins deux ébauches de solution."
      }
    ]
  },

  'ppt-coproj': {
    id: 'ppt-coproj',
    name: 'Conduite de Comité de Projet (COPROJ)',
    code: 'FORM-COPROJ-31',
    totalSlides: 6,
    durationMinutes: 25,
    level: 'Intermédiaire',
    description: "Méthode et bonnes pratiques pour structurer, animer et faire vivre le Comité de Projet opérationnel hebdomadaire avec l'ensemble des responsables de chantiers.",
    slides: [
      {
        slideNumber: 1,
        title: 'Conduite du Comité de Projet (COPROJ)',
        category: 'Périmètre & Mission',
        badge: 'Instance Opérationnelle',
        subtitle: 'Le cœur battant du pilotage hebdomadaire des chantiers',
        bullets: [
          'Le COPROJ réunit chaque semaine le Chef de Projet et les pilotes de chaque chantier.',
          'Durée type : 45 à 60 minutes chrono.',
          'Vocation opérationnelle : lever les blocages du quotidien, coordonner les interfaces inter-équipes et actualiser le plan d\'action.',
          'Différence avec le COPIL : le COPROJ traite du ' + 'comment faire' + ' et résout les problèmes au ras du sol.'
        ],
        keyTakeaway: "Le COPROJ n'est pas un lieu où l'on travaille, mais un lieu où l'on fait le point sur le travail accompli et à venir."
      },
      {
        slideNumber: 2,
        title: 'Ordre du Jour Standard d\'un COPROJ Réussi',
        category: 'Structure',
        badge: 'Timing Rigoureux',
        subtitle: 'Une trame immuable pour une efficacité maximale',
        bullets: [
          '00-05 min : Tour de table météo rapide et validation des décisions de la séance précédente.',
          '05-25 min : Tour d\'avancement des chantiers (1 slide par chantier, faits marquants et livrables).',
          '25-45 min : Traitement approfondi des points de blocage et recherche conjointe de solutions.',
          '45-55 min : Revue et mise à jour du plan d\'action partagé (qui fait quoi pour quand ?).',
          '55-60 min : Clôture et rappel des échéances critiques de la semaine suivante.'
        ],
        keyTakeaway: "Ne laissez jamais un débat technique de détail monopoliser le COPROJ : créez un groupe de travail dédié ad-hoc."
      },
      {
        slideNumber: 3,
        title: 'Suivi d\'Avancement par Chantier & Livrable',
        category: 'Revue Chantiers',
        badge: 'Rigueur de Suivi',
        subtitle: 'Évaluer objectivement la progression sans langue de bois',
        bullets: [
          'Chaque pilote de lot expose : Ce qui a été livré cette semaine, ce qui est en cours, ce qui pose problème.',
          'Indicateur pourcentage d\'avancement physique réel (et non pas temps consommé).',
          'Vérification de la conformité aux exigences du plan qualité.',
          'Identification des adhérences : tel chantier a-t-il besoin d\'un élément fourni par un autre chantier ?'
        ],
        keyTakeaway: "90% d'avancement ne veut rien dire : c'est le dernier 10% qui prend souvent la moitié du temps si les tests sont négligés."
      },
      {
        slideNumber: 4,
        title: 'Animation du Tableau d\'Actions Partagé',
        category: 'Pilotage Actions',
        badge: 'Traçabilité Totale',
        subtitle: 'Chaque engagement pris doit avoir un responsable unique et une date butoir',
        bullets: [
          'Une action sans porteur unique est une action qui ne sera jamais réalisée.',
          'Une action sans échéance précise (ex: ' + 'Fin de semaine 12' + ') est un vœu pieux.',
          'Au début de chaque COPROJ, on passe en revue les actions de la semaine précédente : Fait, En retard, ou Reporté avec justification.',
          'Maintien d\'un taux d\'actions clôturées dans les temps supérieur à 85%.'
        ],
        keyTakeaway: "Le tableau d'actions est la mémoire vivante de l'exécution du projet."
      },
      {
        slideNumber: 5,
        title: 'Traitement des Points Durs & Arbitrages Court Terme',
        category: 'Résolution de Problèmes',
        badge: 'Orientation Solution',
        subtitle: 'Désamorcer les conflits d\'interface entre équipes',
        bullets: [
          'Quand deux équipes ont des visions divergentes, le chef de projet tranche selon le cadre de la charte.',
          'Si le blocage nécessite des moyens financiers ou humains supplémentaires : formaliser une fiche d\'arbitrage pour le COPIL.',
          'Encourager la franchise : il n\'y a aucune honte à signaler un retard si on le signale assez tôt pour réagir.'
        ],
        keyTakeaway: "Le courage managérial du chef de projet se révèle dans sa capacité à trancher les désaccords opérationnels."
      },
      {
        slideNumber: 6,
        title: 'Compte-Rendu Flash & Clôture',
        category: 'Restitution',
        badge: 'Diffusion Rapide',
        subtitle: 'Garder une trace formelle sans surcharge administrative',
        bullets: [
          'Le compte-rendu du COPROJ doit être diffusé dans les 2 heures suivant la réunion.',
          'Format ultra synthétique : 1 tableau d\'actions mis à jour + liste des 3 décisions clés.',
          'Permet à ceux qui étaient absents d\'être immédiatement alignés.',
          'Alimente directement la rédaction du rapport flash de fin de semaine.'
        ],
        keyTakeaway: "Un compte-rendu envoyé 5 jours plus tard ne sert plus à rien : la vitesse de diffusion conditionne l'action."
      }
    ]
  },

  'ppt-copil': {
    id: 'ppt-copil',
    name: 'Conduite de Comité de Pilotage (COPIL)',
    code: 'FORM-COPIL-32',
    totalSlides: 6,
    durationMinutes: 30,
    level: 'Gouvernance',
    description: "Formation managériale pour préparer et présider le Comité de Pilotage stratégique : posture, arbitrage budgétaire, passage de jalons et prise de décision au plus haut niveau.",
    slides: [
      {
        slideNumber: 1,
        title: 'Gouvernance Stratégique : Le Comité de Pilotage (COPIL)',
        category: 'Rôle & Missions',
        badge: 'Pouvoir Souverain',
        subtitle: 'L\'instance de décision ultime qui engage la responsabilité de l\'entreprise',
        bullets: [
          'Le COPIL réunit le Sponsor, la Direction Générale, les Directeurs Métiers et le Chef de Projet.',
          'Fréquence mensuelle ou aux grands jalons clés de passage (Gate Reviews).',
          'Pouvoirs exclusifs : valider le budget, autoriser le passage d\'un jalon, arbitrer les demandes de modification majeures, décider de l\'arrêt ou de la poursuite du projet.',
          'Le chef de projet y joue le rôle d\'instructeur et de rapporteur stratégique.'
        ],
        keyTakeaway: "Le COPIL ne fait pas d'opérationnel : il tranche les orientations majeures et garantit l'alignement sur la stratégie d'entreprise."
      },
      {
        slideNumber: 2,
        title: 'Dossier de Présentation : Règle des 3 Niveaux de Lecture',
        category: 'Communication Exécutive',
        badge: 'Impact Managérial',
        subtitle: 'Captiver et rassurer la direction en respectant son niveau d\'exigence',
        bullets: [
          'Niveau 1 (Executive Summary) : La synthèse globale sur une diapositive (Météo, Jalons, Budget, Alerte n°1).',
          'Niveau 2 (Détail par domaine) : Suivi analytique de l\'avancement, indicateurs financiers VA/VP/CR.',
          'Niveau 3 (Dossiers d\'arbitrage) : Les fiches de décision soumises au vote des décideurs.',
          'Posture du chef de projet : Maîtrise, franchise, neutralité et force de proposition.'
        ],
        keyTakeaway: "Ne commencez jamais un COPIL par des détails techniques : commencez par le résultat financier et calendaire global."
      },
      {
        slideNumber: 3,
        title: 'Analyse de la Valeur Acquise (VA / VP / CR) & Performance',
        category: 'Indicateurs Financiers',
        badge: 'Earned Value',
        subtitle: 'Mesurer la rentabilité réelle et la trajectoire budgétaire finale',
        bullets: [
          'VP (Valeur Planifiée) : Le budget que l\'on avait prévu d\'avoir dépensé à cette date.',
          'CR (Coût Réel) : Les dépenses effectives réellement engagées et facturées.',
          'VA (Valeur Acquise) : La valeur financière réelle du travail effectivement achevé.',
          'CPI (Cost Performance Index) = VA / CR : Si CPI < 1, le projet dépense plus que prévu pour ce qu\'il produit !',
          'SPI (Schedule Performance Index) = VA / VP : Si SPI < 1, le projet avance moins vite que prévu.'
        ],
        keyTakeaway: "Le CPI et le SPI sont les indicateurs souverains les plus respectés par les directions financières internationales."
      },
      {
        slideNumber: 4,
        title: 'Passage de Jalon (Gate Review) & Autorisation d\'Étape',
        category: 'Portes de Passage',
        badge: 'Go / No-Go',
        subtitle: 'Valider la complétude d\'une phase avant d\'autoriser l\'engagement des dépenses suivantes',
        bullets: [
          'Chaque jalon (J1, J2, J3, J4) constitue une porte d\'audit formelle (Gate Review).',
          'Le COPIL vérifie la grille de critères obligatoires (ex: 100% des spécifications signées pour J1).',
          'Trois issues possibles : GO sans réserve (passage accordé), GO sous conditions (réserves à lever sous 15 jours), NO-GO (gel de la phase jusqu\'à mise en conformité).',
          'Signature formelle du Procès-Verbal de passage de jalon.'
        ],
        keyTakeaway: "Ne jamais passer une porte de jalon en catimini : cela transfère la responsabilité d'un échec futur sur les épaules du chef de projet."
      },
      {
        slideNumber: 5,
        title: 'Arbitrage des Demandes de Changement & Modifications Budgétaires',
        category: 'Arbitrages',
        badge: 'Impact Financier',
        subtitle: 'Formaliser les avenants contractuels et les ajustements d\'enveloppe',
        bullets: [
          'Présentation des demandes de modification ayant un impact sur les coûts ou les délais.',
          'Exposé comparatif : Gain attendu vs Coût additionnel vs Risque d\'échec.',
          'Le Sponsor valide l\'avenant budgétaire ou arbitre le rejet de la demande.',
          'Consignation immédiate de la décision au Registre officiel des Décisions du COPIL.'
        ],
        keyTakeaway: "Le COPIL a la légitimité financière pour allouer du budget supplémentaire : le chef de projet ne doit pas hésiter à solliciter ses arbitrages."
      },
      {
        slideNumber: 6,
        title: 'Rédaction du Procès-Verbal & Traçabilité Juridique',
        category: 'Formalisation',
        badge: 'Acte Officiel',
        subtitle: 'Consigner les engagements mutuels avec valeur contractuelle',
        bullets: [
          'Le PV de réunion du COPIL est un document à haute valeur probante.',
          'Il consigne nominativement les membres présents, les votes d\'arbitrage et les mandats confiés.',
          'Généré et imprimable directement au format PDF conforme depuis votre outil FAAB PROJET.',
          'Signé par le Président du COPIL et le Chef de Projet pour archivage pérenne.'
        ],
        keyTakeaway: "Les paroles s'envolent, les écrits restent : le PV de COPIL est l'assurance tous risques du chef de projet."
      }
    ]
  }
};

interface ProjectTrainingViewerProps {
  initialCourseId?: string;
  onClose?: () => void;
}

export const ProjectTrainingViewer: React.FC<ProjectTrainingViewerProps> = ({
  initialCourseId = 'ppt-lancement'
}) => {
  const [selectedCourseId, setSelectedCourseId] = useState<string>(
    trainingCourses[initialCourseId] ? initialCourseId : 'ppt-lancement'
  );
  const [currentSlideIndex, setCurrentSlideIndex] = useState<number>(0);
  const [showThumbnails, setShowThumbnails] = useState<boolean>(true);

  const activeCourse = trainingCourses[selectedCourseId] || trainingCourses['ppt-lancement'];
  const totalSlides = activeCourse.slides.length;
  const currentSlide = activeCourse.slides[currentSlideIndex] || activeCourse.slides[0];

  const handleSelectCourse = (id: string) => {
    setSelectedCourseId(id);
    setCurrentSlideIndex(0);
  };

  const handlePrevSlide = () => {
    if (currentSlideIndex > 0) {
      setCurrentSlideIndex(currentSlideIndex - 1);
    }
  };

  const handleNextSlide = () => {
    if (currentSlideIndex < totalSlides - 1) {
      setCurrentSlideIndex(currentSlideIndex + 1);
    }
  };

  const handlePrintCourse = () => {
    window.print();
  };

  return (
    <div className="training-viewer-container" style={{
      display: 'flex',
      flexDirection: 'column',
      height: '100%',
      background: '#090d16',
      color: '#f8fafc',
      overflow: 'hidden'
    }}>
      {/* Top Bar / Navigation (Hidden on print) */}
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
        {/* Left: Course Selector */}
        <div style={{ display: 'flex', alignItems: 'center', gap: '0.8rem' }}>
          <span style={{ fontSize: '1.4rem' }}>🎓</span>
          <div>
            <div style={{ fontSize: '0.72rem', color: '#f59e0b', textTransform: 'uppercase', fontWeight: 700, letterSpacing: '0.5px' }}>
              Supports de Formation Projet & Pédagogie
            </div>
            <select
              value={selectedCourseId}
              onChange={(e) => handleSelectCourse(e.target.value)}
              style={{
                background: 'rgba(30, 41, 59, 0.9)',
                border: '1px solid rgba(245, 158, 11, 0.4)',
                color: '#fbbf24',
                padding: '5px 12px',
                borderRadius: '6px',
                fontWeight: 700,
                fontSize: '0.9rem',
                cursor: 'pointer'
              }}
            >
              {Object.values(trainingCourses).map(c => (
                <option key={c.id} value={c.id} style={{ background: '#0f172a', color: '#f8fafc' }}>
                  {c.name} ({c.code} • {c.totalSlides} diapositives)
                </option>
              ))}
            </select>
          </div>
        </div>

        {/* Center: Slide Navigator */}
        <div style={{ display: 'flex', alignItems: 'center', gap: '1rem' }}>
          <button
            onClick={handlePrevSlide}
            disabled={currentSlideIndex === 0}
            style={{
              background: currentSlideIndex === 0 ? 'rgba(255,255,255,0.04)' : '#334155',
              border: 'none',
              color: currentSlideIndex === 0 ? '#64748b' : '#fff',
              padding: '6px 14px',
              borderRadius: '6px',
              fontSize: '0.85rem',
              fontWeight: 600,
              cursor: currentSlideIndex === 0 ? 'not-allowed' : 'pointer',
              display: 'flex',
              alignItems: 'center',
              gap: '4px'
            }}
          >
            ◀ Précédent
          </button>

          <div style={{
            fontSize: '0.85rem',
            fontWeight: 700,
            color: '#cbd5e1',
            background: 'rgba(255,255,255,0.06)',
            padding: '4px 12px',
            borderRadius: '20px'
          }}>
            Diapositive <span style={{ color: '#fbbf24' }}>{currentSlideIndex + 1}</span> / {totalSlides}
          </div>

          <button
            onClick={handleNextSlide}
            disabled={currentSlideIndex === totalSlides - 1}
            style={{
              background: currentSlideIndex === totalSlides - 1 ? 'rgba(255,255,255,0.04)' : 'linear-gradient(135deg, #f59e0b, #d97706)',
              border: 'none',
              color: currentSlideIndex === totalSlides - 1 ? '#64748b' : '#fff',
              padding: '6px 14px',
              borderRadius: '6px',
              fontSize: '0.85rem',
              fontWeight: 600,
              cursor: currentSlideIndex === totalSlides - 1 ? 'not-allowed' : 'pointer',
              display: 'flex',
              alignItems: 'center',
              gap: '4px',
              boxShadow: currentSlideIndex === totalSlides - 1 ? 'none' : '0 4px 12px rgba(245, 158, 11, 0.3)'
            }}
          >
            Suivant ▶
          </button>
        </div>

        {/* Right: Actions */}
        <div style={{ display: 'flex', alignItems: 'center', gap: '0.8rem' }}>
          <button
            onClick={() => setShowThumbnails(!showThumbnails)}
            style={{
              background: showThumbnails ? 'rgba(245, 158, 11, 0.2)' : 'rgba(255,255,255,0.08)',
              border: `1px solid ${showThumbnails ? '#f59e0b' : 'rgba(255,255,255,0.15)'}`,
              color: showThumbnails ? '#f59e0b' : '#cbd5e1',
              padding: '6px 12px',
              borderRadius: '6px',
              fontSize: '0.8rem',
              fontWeight: 600,
              cursor: 'pointer'
            }}
          >
            📑 {showThumbnails ? "Masquer Plan" : "Afficher Plan"}
          </button>

          <button
            onClick={handlePrintCourse}
            style={{
              background: 'linear-gradient(135deg, #0284c7, #2563eb)',
              border: 'none',
              color: '#fff',
              padding: '6px 14px',
              borderRadius: '6px',
              fontSize: '0.82rem',
              fontWeight: 700,
              cursor: 'pointer',
              display: 'flex',
              alignItems: 'center',
              gap: '6px'
            }}
          >
            📥 Imprimer le Support
          </button>
        </div>
      </div>

      {/* Main Slide Workspace */}
      <div style={{ flex: 1, display: 'flex', overflow: 'hidden' }}>
        {/* Left Slide Plan / Thumbnails (Optional) */}
        {showThumbnails && (
          <div className="no-print" style={{
            width: '280px',
            background: 'rgba(15, 23, 42, 0.9)',
            borderRight: '1px solid rgba(255, 255, 255, 0.08)',
            overflowY: 'auto',
            padding: '1rem'
          }}>
            <div style={{ fontSize: '0.78rem', color: '#94a3b8', textTransform: 'uppercase', fontWeight: 700, marginBottom: '0.8rem' }}>
              Index des Diapositives ({totalSlides})
            </div>
            <div style={{ display: 'flex', flexDirection: 'column', gap: '0.5rem' }}>
              {activeCourse.slides.map((s, idx) => {
                const isActive = idx === currentSlideIndex;
                return (
                  <button
                    key={s.slideNumber}
                    onClick={() => setCurrentSlideIndex(idx)}
                    style={{
                      background: isActive ? 'linear-gradient(135deg, rgba(245, 158, 11, 0.2), rgba(217, 119, 6, 0.15))' : 'rgba(255, 255, 255, 0.03)',
                      border: `1px solid ${isActive ? '#f59e0b' : 'rgba(255, 255, 255, 0.05)'}`,
                      borderRadius: '8px',
                      padding: '8px 10px',
                      textAlign: 'left',
                      cursor: 'pointer',
                      display: 'flex',
                      flexDirection: 'column',
                      gap: '3px',
                      transition: 'all 0.15s ease'
                    }}
                  >
                    <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between' }}>
                      <span style={{ fontSize: '0.7rem', fontWeight: 800, color: isActive ? '#f59e0b' : '#64748b' }}>
                        SLIDE {s.slideNumber < 10 ? `0${s.slideNumber}` : s.slideNumber}
                      </span>
                      {s.badge && (
                        <span style={{ fontSize: '0.65rem', background: 'rgba(255,255,255,0.08)', padding: '1px 6px', borderRadius: '4px', color: '#cbd5e1' }}>
                          {s.badge}
                        </span>
                      )}
                    </div>
                    <div style={{ fontSize: '0.8rem', fontWeight: 600, color: isActive ? '#fff' : '#cbd5e1', lineHeight: 1.3 }}>
                      {s.title}
                    </div>
                  </button>
                );
              })}
            </div>
          </div>
        )}

        {/* Center: Slide Stage */}
        <div style={{
          flex: 1,
          overflowY: 'auto',
          padding: '2.5rem',
          display: 'flex',
          justifyContent: 'center',
          alignItems: 'center',
          background: 'radial-gradient(ellipse at top, #1e293b 0%, #090d16 100%)'
        }}>
          {/* Main Slide Card 16:9 Presentation Format */}
          <div className="slide-card" style={{
            width: '100%',
            maxWidth: '1020px',
            minHeight: '600px',
            background: 'linear-gradient(145deg, #1e293b 0%, #0f172a 100%)',
            border: '1px solid rgba(245, 158, 11, 0.25)',
            borderRadius: '16px',
            boxShadow: '0 20px 50px rgba(0,0,0,0.6)',
            padding: '40px 48px',
            display: 'flex',
            flexDirection: 'column',
            justifyContent: 'space-between',
            position: 'relative',
            overflow: 'hidden'
          }}>
            {/* Top Slide Header */}
            <div>
              <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginBottom: '16px' }}>
                <div style={{ display: 'flex', alignItems: 'center', gap: '10px' }}>
                  <span style={{
                    background: 'rgba(245, 158, 11, 0.15)',
                    color: '#f59e0b',
                    border: '1px solid rgba(245, 158, 11, 0.3)',
                    padding: '3px 10px',
                    borderRadius: '20px',
                    fontSize: '0.75rem',
                    fontWeight: 700,
                    textTransform: 'uppercase',
                    letterSpacing: '0.5px'
                  }}>
                    {currentSlide.category}
                  </span>
                  {currentSlide.badge && (
                    <span style={{
                      background: 'rgba(56, 189, 248, 0.15)',
                      color: '#38bdf8',
                      border: '1px solid rgba(56, 189, 248, 0.3)',
                      padding: '3px 10px',
                      borderRadius: '20px',
                      fontSize: '0.75rem',
                      fontWeight: 700
                    }}>
                      {currentSlide.badge}
                    </span>
                  )}
                </div>

                <div style={{ fontSize: '0.75rem', color: '#64748b', fontWeight: 600 }}>
                  {activeCourse.code} • Diapositive {currentSlide.slideNumber} sur {totalSlides}
                </div>
              </div>

              {/* Slide Title */}
              <h2 style={{
                margin: '0 0 8px 0',
                fontSize: '1.85rem',
                fontWeight: 800,
                color: '#f8fafc',
                lineHeight: 1.25,
                letterSpacing: '-0.5px'
              }}>
                {currentSlide.title}
              </h2>

              {currentSlide.subtitle && (
                <p style={{
                  margin: '0 0 24px 0',
                  fontSize: '1.05rem',
                  color: '#94a3b8',
                  lineHeight: 1.4
                }}>
                  {currentSlide.subtitle}
                </p>
              )}
            </div>

            {/* Slide Body: Bullet Points */}
            <div style={{ margin: '15px 0', flex: 1, display: 'flex', flexDirection: 'column', gap: '14px' }}>
              {currentSlide.bullets.map((bullet, bIdx) => (
                <div key={bIdx} style={{ display: 'flex', alignItems: 'flex-start', gap: '12px' }}>
                  <span style={{
                    color: '#f59e0b',
                    fontSize: '1.1rem',
                    lineHeight: '1.4'
                  }}>
                    ✦
                  </span>
                  <div style={{
                    fontSize: '1.02rem',
                    lineHeight: 1.55,
                    color: '#e2e8f0',
                    fontWeight: 400
                  }}>
                    {bullet}
                  </div>
                </div>
              ))}

              {/* Structured Metrics if present */}
              {currentSlide.structuredData && (
                <div style={{
                  display: 'grid',
                  gridTemplateColumns: `repeat(${currentSlide.structuredData.length}, 1fr)`,
                  gap: '14px',
                  marginTop: '16px'
                }}>
                  {currentSlide.structuredData.map((d, dIdx) => (
                    <div key={dIdx} style={{
                      background: 'rgba(255, 255, 255, 0.04)',
                      border: '1px solid rgba(255, 255, 255, 0.08)',
                      borderRadius: '10px',
                      padding: '12px 16px',
                      textAlign: 'center'
                    }}>
                      <div style={{ fontSize: '0.75rem', color: '#94a3b8', textTransform: 'uppercase', fontWeight: 600, marginBottom: '4px' }}>
                        {d.label}
                      </div>
                      <div style={{ fontSize: '1.15rem', fontWeight: 800, color: d.color || '#38bdf8' }}>
                        {d.value}
                      </div>
                    </div>
                  ))}
                </div>
              )}

              {/* Diagram Steps if present */}
              {currentSlide.diagramSteps && (
                <div style={{
                  display: 'grid',
                  gridTemplateColumns: `repeat(${currentSlide.diagramSteps.length}, 1fr)`,
                  gap: '12px',
                  marginTop: '18px'
                }}>
                  {currentSlide.diagramSteps.map((step) => (
                    <div key={step.step} style={{
                      background: 'rgba(255, 255, 255, 0.03)',
                      border: '1px solid rgba(255, 255, 255, 0.06)',
                      borderRadius: '10px',
                      padding: '14px',
                      display: 'flex',
                      flexDirection: 'column',
                      alignItems: 'center',
                      textAlign: 'center'
                    }}>
                      <span style={{ fontSize: '1.8rem', marginBottom: '6px' }}>{step.icon}</span>
                      <strong style={{ fontSize: '0.88rem', color: '#f8fafc', marginBottom: '4px' }}>{step.title}</strong>
                      <span style={{ fontSize: '0.74rem', color: '#94a3b8' }}>{step.desc}</span>
                    </div>
                  ))}
                </div>
              )}
            </div>

            {/* Bottom Key Takeaway Callout */}
            {currentSlide.keyTakeaway && (
              <div style={{
                marginTop: '20px',
                background: 'linear-gradient(135deg, rgba(245, 158, 11, 0.1) 0%, rgba(217, 119, 6, 0.05) 100%)',
                borderLeft: '4px solid #f59e0b',
                borderRadius: '0 10px 10px 0',
                padding: '14px 20px',
                display: 'flex',
                alignItems: 'center',
                gap: '12px'
              }}>
                <span style={{ fontSize: '1.4rem' }}>💡</span>
                <div>
                  <div style={{ fontSize: '0.72rem', fontWeight: 800, color: '#fbbf24', textTransform: 'uppercase', letterSpacing: '0.5px' }}>
                    À retenir pour le Chef de Projet
                  </div>
                  <div style={{ fontSize: '0.92rem', color: '#fef3c7', fontWeight: 500, fontStyle: 'italic' }}>
                    "{currentSlide.keyTakeaway}"
                  </div>
                </div>
              </div>
            )}

            {/* Slide Footer */}
            <div style={{
              marginTop: '22px',
              paddingTop: '14px',
              borderTop: '1px solid rgba(255, 255, 255, 0.08)',
              display: 'flex',
              alignItems: 'center',
              justifyContent: 'space-between',
              fontSize: '0.72rem',
              color: '#64748b'
            }}>
              <div>
                FAAB PROJET • Support Pédagogique Officiel • {activeCourse.name}
              </div>
              <div style={{ display: 'flex', gap: '6px' }}>
                {activeCourse.slides.map((_, i) => (
                  <div
                    key={i}
                    onClick={() => setCurrentSlideIndex(i)}
                    style={{
                      width: i === currentSlideIndex ? '20px' : '6px',
                      height: '6px',
                      borderRadius: '3px',
                      background: i === currentSlideIndex ? '#f59e0b' : 'rgba(255,255,255,0.2)',
                      cursor: 'pointer',
                      transition: 'all 0.2s ease'
                    }}
                  />
                ))}
              </div>
            </div>
          </div>
        </div>
      </div>
    </div>
  );
};
