# Présentation : Datacenter & Infrastructure Numérique de Santé en Côte d'Ivoire

Présentation institutionnelle du **Datacenter National de Santé de Côte d'Ivoire**, développé sous l'égide du **Ministère de la Santé, de l'Hygiène Publique et de la Couverture Maladie Universelle (MSHPCMU)** et de la **Direction de l'Informatique et de la Santé Digitale (DISD)**.

🌐 **Lien en ligne (Démo Web interactive)** : [https://cyrhinov-bit.github.io/datacenter/](https://cyrhinov-bit.github.io/datacenter/)

---

## 📋 Contenu du Projet

1. **`index.html`** :
   - Application web interactive de présentation (HTML5 / Tailwind CSS / Lucide Icons / Canvas Particles).
   - Modes de navigation :
     - Navigation manuelle (flèches clavier, boutons de navigation).
     - Défilement automatique (*Lecture Auto*) avec sélecteur de vitesse (**1.5s**, **2.5s**, **4.0s**).
     - Mode plein écran (<kbd>F</kbd>).
     - Révélation progressive étape par étape des éléments.
     - Intégration des photos réelles de l'infrastructure Datacenter.

2. **Générateurs PowerPoint (`generate_pptx_style.py`)** :
   - Script Python (`python-pptx`) générant la présentation 16:9 au format `.pptx`.
   - Respect strict de la charte visuelle moderne (capsules, badges ronds, palette institutionnelle).
   - Intégration haute résolution des photos et logos officiels.

3. **`assets/`** :
   - Photos techniques réelles :
     - Racks et allées de confinement (`datacenter1.png`)
     - Supervision et monitoring EcoStruxure (`datacenter2.png`)
     - Aéroréfrigérants externes & climatisation (`datacenter3.png`)
     - Groupes électrogènes industriels 165 kVA (`datacenter4.png`)
     - Salle IT et sécurité incendie gaz inerte (`datacenter_slide_11_image_3.png`)
   - Logos officiels (Ministère de la Santé, DISD).

---

## 🚀 Utilisation

### 1. Visualiser la présentation Web interactive
Ouvrez simplement le fichier `index.html` dans n'importe quel navigateur moderne (Chrome, Edge, Firefox, Safari).

### 2. Régénérer le fichier PowerPoint (.pptx)
Installez les dépendances requises :
```bash
pip install python-pptx pillow lxml
```

Générez la présentation :
```bash
python generate_pptx_style.py
```
Le fichier `Presentation_Datacenter_Sante_CI_Style.pptx` sera généré à la racine.

---

## 🏛️ Crédits & Contexte Institutionnel
- **Ministère de la Santé, de l'Hygiène Publique et de la Couverture Maladie Universelle (MSHPCMU)**
- **Direction de l'Informatique et de la Santé Digitale (DISD)**
- **Présentateur** : M. OUATTARA Yacouba, Sous-Directeur des Infrastructures

