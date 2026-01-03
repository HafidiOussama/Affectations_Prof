class ApplicationAffectation {
    constructor() {
        this.professeurs = [];
        this.salles = [];
        this.matieres = [];
        this.affectations = [];
        this.dates_examens = [];
        this.type_examen = "امتحان عادي";
        this.selectedProfessorIndex = null;
        this.selectedSubjects = new Set();
        
        this.init();
    }

    init() {
        this.bindEvents();
        this.loadFromLocalStorage();
        this.updateStats();
        this.updateStatus("جاهز • نظام توزيع الأساتذة");
    }

    bindEvents() {
        // Navigation
        document.querySelectorAll('.nav-tab').forEach(tab => {
            tab.addEventListener('click', (e) => this.switchTab(e.target.dataset.tab));
        });

        document.getElementById('mobileMenuBtn').addEventListener('click', () => {
            document.querySelector('.nav-tabs').classList.toggle('active');
        });

        // Boutons Accueil
        document.getElementById('startConfig').addEventListener('click', () => {
            this.switchTab('professeurs');
        });

        // Boutons Professeurs
        document.getElementById('importExcelBtn').addEventListener('click', () => this.importerExcel());
        document.getElementById('addProfessorBtn').addEventListener('click', () => this.showProfessorModal());
        document.getElementById('editProfessorBtn').addEventListener('click', () => this.modifierProfesseur());
        document.getElementById('deleteProfessorBtn').addEventListener('click', () => this.supprimerProfesseur());
        document.getElementById('manageUnavailabilityBtn').addEventListener('click', () => this.gererIndisponibilites());
        document.getElementById('deleteUnavailabilityBtn').addEventListener('click', () => this.supprimerIndisponibilite());
        document.getElementById('extractSubjectsBtn').addEventListener('click', () => this.extraireMatieres());

        // Boutons Paramètres
        document.getElementById('generateSubjectsForm').addEventListener('click', () => this.genererFormulaireMatieres());
        document.getElementById('fillFromProfessors').addEventListener('click', () => this.remplirMatieresDepuisProfesseurs());
        document.getElementById('deleteSubjectBtn').addEventListener('click', () => this.supprimerMatiere());

        // Boutons Affectations
        document.getElementById('generateAssignmentsBtn').addEventListener('click', () => this.genererAffectations());
        document.getElementById('exportExcelBtn').addEventListener('click', () => this.showExcelConfigModal());

        // Modal Professeur
        document.getElementById('saveProfessorBtn').addEventListener('click', () => this.saveProfessor());
        document.querySelectorAll('#professorModal .close-btn').forEach(btn => {
            btn.addEventListener('click', () => this.closeModal('professorModal'));
        });

        // Modal Indisponibilités
        document.getElementById('saveUnavailabilityBtn').addEventListener('click', () => this.saveUnavailability());
        document.querySelectorAll('#unavailabilityModal .close-btn').forEach(btn => {
            btn.addEventListener('click', () => this.closeModal('unavailabilityModal'));
        });

        // Modal Matières
        document.getElementById('useSelectedSubjectsBtn').addEventListener('click', () => this.utiliserMatieresSelectionnees());
        document.getElementById('deleteSelectedSubjectsBtn').addEventListener('click', () => this.supprimerMatieresSelectionnees());
        document.querySelectorAll('#subjectsExtractModal .close-btn').forEach(btn => {
            btn.addEventListener('click', () => this.closeModal('subjectsExtractModal'));
        });

        // Modal Excel
        document.getElementById('generateExcelBtn').addEventListener('click', () => this.genererExcel());
        document.querySelectorAll('#excelConfigModal .close-btn').forEach(btn => {
            btn.addEventListener('click', () => this.closeModal('excelConfigModal'));
        });

        // Fermer les modals en cliquant en dehors
        document.querySelectorAll('.modal').forEach(modal => {
            modal.addEventListener('click', (e) => {
                if (e.target === modal) {
                    this.closeModal(modal.id);
                }
            });
        });

        // Sélection de ligne dans le tableau des professeurs
        document.addEventListener('click', (e) => {
            const row = e.target.closest('tr[data-index]');
            if (row && e.target.closest('#professorsTableBody')) {
                this.selectProfessorRow(row);
            }
        });
    }

    // ========== GESTION DE L'INTERFACE ==========

    switchTab(tabName) {
        // Cacher tous les onglets
        document.querySelectorAll('.tab-content').forEach(tab => {
            tab.classList.remove('active');
        });
        
        // Désactiver tous les boutons d'onglets
        document.querySelectorAll('.nav-tab').forEach(tab => {
            tab.classList.remove('active');
        });
        
        // Afficher l'onglet sélectionné
        document.getElementById(tabName).classList.add('active');
        
        // Activer le bouton correspondant
        document.querySelector(`[data-tab="${tabName}"]`).classList.add('active');
        
        // Fermer le menu mobile si ouvert
        document.querySelector('.nav-tabs').classList.remove('active');
        
        this.updateStatus(`تم التبديل إلى ${this.getTabName(tabName)}`);
    }

    getTabName(tabId) {
        const names = {
            'accueil': 'الصفحة الرئيسية',
            'professeurs': 'الأساتذة',
            'parametres': 'الإعدادات',
            'affectations': 'التوزيع'
        };
        return names[tabId] || tabId;
    }

    showModal(modalId) {
        document.getElementById(modalId).classList.add('active');
    }

    closeModal(modalId) {
        document.getElementById(modalId).classList.remove('active');
    }

    updateStatus(message) {
        document.getElementById('statusMessage').textContent = message;
    }

    // ========== GESTION DES DONNÉES ==========

    saveToLocalStorage() {
        const data = {
            professeurs: this.professeurs,
            matieres: this.matieres,
            affectations: this.affectations,
            type_examen: this.type_examen
        };
        localStorage.setItem('affectation_app', JSON.stringify(data));
    }

    loadFromLocalStorage() {
        const data = JSON.parse(localStorage.getItem('affectation_app'));
        if (data) {
            this.professeurs = data.professeurs || [];
            this.matieres = data.matieres || [];
            this.affectations = data.affectations || [];
            this.type_examen = data.type_examen || "امتحان عادي";
            
            this.afficherProfesseurs();
            this.updateStats();
        }
    }

    // ========== GESTION DES PROFESSEURS ==========

    afficherProfesseurs() {
        const tbody = document.getElementById('professorsTableBody');
        tbody.innerHTML = '';
        
        this.professeurs.forEach((prof, index) => {
            const indispoText = prof.indisponibilites && prof.indisponibilites.length > 0 
                ? prof.indisponibilites.map(([jour, periode]) => `${jour}-${periode === 'matin' ? 'صباح' : 'مساء'}`).join(', ')
                : 'لا يوجد';
            
            const statut = prof.indisponibilites && prof.indisponibilites.length > 0 
                ? '<span class="status-badge status-unavailable">غير متاح</span>'
                : '<span class="status-badge status-available">متاح</span>';
            
            const tr = document.createElement('tr');
            tr.dataset.index = index;
            tr.innerHTML = `
                <td>${prof.nom}</td>
                <td>${prof.matiere}</td>
                <td>${prof.numero || ''}</td>
                <td>${indispoText}</td>
                <td>${statut}</td>
            `;
            tbody.appendChild(tr);
        });
    }

    selectProfessorRow(row) {
        // Enlever la sélection précédente
        document.querySelectorAll('#professorsTableBody tr').forEach(r => {
            r.classList.remove('selected-row');
        });
        
        // Ajouter la sélection à la nouvelle ligne
        row.classList.add('selected-row');
        this.selectedProfessorIndex = parseInt(row.dataset.index);
    }

    showProfessorModal(professor = null) {
        const modal = document.getElementById('professorModal');
        const title = document.getElementById('modalTitle');
        const form = document.getElementById('professorForm');
        
        if (professor) {
            title.textContent = 'تعديل الأستاذ';
            document.getElementById('fullName').value = professor.nom || '';
            document.getElementById('subject').value = professor.matiere || '';
            document.getElementById('rentalNumber').value = professor.numero || '';
        } else {
            title.textContent = 'إضافة أستاذ جديد';
            form.reset();
        }
        
        this.showModal('professorModal');
    }

    saveProfessor() {
        const nom = document.getElementById('fullName').value.trim();
        const matiere = document.getElementById('subject').value.trim();
        const numero = document.getElementById('rentalNumber').value.trim();
        
        if (!nom || !matiere) {
            Swal.fire('خطأ', 'يرجى ملء جميع الحقول الإلزامية', 'error');
            return;
        }
        
        const professeur = {
            nom: nom,
            matiere: matiere,
            numero: numero,
            indisponibilites: []
        };
        
        if (this.selectedProfessorIndex !== null && document.getElementById('modalTitle').textContent === 'تعديل الأستاذ') {
            // Mettre à jour le professeur existant
            professeur.indisponibilites = this.professeurs[this.selectedProfessorIndex].indisponibilites || [];
            this.professeurs[this.selectedProfessorIndex] = professeur;
        } else {
            // Ajouter un nouveau professeur
            this.professeurs.push(professeur);
        }
        
        this.afficherProfesseurs();
        this.updateStats();
        this.saveToLocalStorage();
        this.closeModal('professorModal');
        this.selectedProfessorIndex = null;
        
        Swal.fire('نجاح', 'تم حفظ الأستاذ بنجاح', 'success');
    }

    modifierProfesseur() {
        if (this.selectedProfessorIndex === null) {
            Swal.fire('تنبيه', 'يرجى اختيار أستاذ أولاً', 'warning');
            return;
        }
        
        const professeur = this.professeurs[this.selectedProfessorIndex];
        this.showProfessorModal(professeur);
    }

    supprimerProfesseur() {
        if (this.selectedProfessorIndex === null) {
            Swal.fire('تنبيه', 'يرجى اختيار أستاذ أولاً', 'warning');
            return;
        }
        
        Swal.fire({
            title: 'تأكيد الحذف',
            text: 'هل أنت متأكد من حذف هذا الأستاذ؟',
            icon: 'warning',
            showCancelButton: true,
            confirmButtonText: 'نعم، احذف',
            cancelButtonText: 'إلغاء',
            confirmButtonColor: '#d33'
        }).then((result) => {
            if (result.isConfirmed) {
                this.professeurs.splice(this.selectedProfessorIndex, 1);
                this.afficherProfesseurs();
                this.updateStats();
                this.saveToLocalStorage();
                this.selectedProfessorIndex = null;
                
                Swal.fire('نجاح', 'تم حذف الأستاذ بنجاح', 'success');
            }
        });
    }

    gererIndisponibilites() {
        if (this.selectedProfessorIndex === null) {
            Swal.fire('تنبيه', 'يرجى اختيار أستاذ أولاً', 'warning');
            return;
        }
        
        const professeur = this.professeurs[this.selectedProfessorIndex];
        document.getElementById('unavailabilityTitle').textContent = `إدارة عدم التوفر - ${professeur.nom}`;
        
        // Remplir les champs existants
        if (professeur.indisponibilites && professeur.indisponibilites.length > 0) {
            const jours = [...new Set(professeur.indisponibilites.map(([jour]) => jour))];
            document.getElementById('unavailabilityDays').value = jours.join(', ');
            
            const periodes = [...new Set(professeur.indisponibilites.map(([, periode]) => periode))];
            document.getElementById('morningPeriod').checked = periodes.includes('matin');
            document.getElementById('eveningPeriod').checked = periodes.includes('soir');
        } else {
            document.getElementById('unavailabilityDays').value = '';
            document.getElementById('morningPeriod').checked = false;
            document.getElementById('eveningPeriod').checked = false;
        }
        
        this.showModal('unavailabilityModal');
    }

    saveUnavailability() {
        const joursText = document.getElementById('unavailabilityDays').value.trim();
        const matin = document.getElementById('morningPeriod').checked;
        const soir = document.getElementById('eveningPeriod').checked;
        
        if (!joursText) {
            Swal.fire('خطأ', 'يرجى إدخال أيام عدم التوفر', 'error');
            return;
        }
        
        if (!matin && !soir) {
            Swal.fire('خطأ', 'يرجى اختيار فترة عدم التوفر على الأقل', 'error');
            return;
        }
        
        const jours = joursText.split(',').map(j => j.trim()).filter(j => j);
        const indisponibilites = [];
        
        jours.forEach(jour => {
            if (matin) indisponibilites.push([jour, 'matin']);
            if (soir) indisponibilites.push([jour, 'soir']);
        });
        
        this.professeurs[this.selectedProfessorIndex].indisponibilites = indisponibilites;
        this.afficherProfesseurs();
        this.updateStats();
        this.saveToLocalStorage();
        this.closeModal('unavailabilityModal');
        
        Swal.fire('نجاح', 'تم تحديث عدم التوفر بنجاح', 'success');
    }

    supprimerIndisponibilite() {
        if (this.selectedProfessorIndex === null) {
            Swal.fire('تنبيه', 'يرجى اختيار أستاذ أولاً', 'warning');
            return;
        }
        
        const professeur = this.professeurs[this.selectedProfessorIndex];
        
        if (!professeur.indisponibilites || professeur.indisponibilites.length === 0) {
            Swal.fire('تنبيه', 'لا توجد فترات عدم توفر لحذفها', 'info');
            return;
        }
        
        Swal.fire({
            title: 'حذف عدم التوفر',
            html: `
                <div style="text-align: right; margin: 20px 0;">
                    <p>اختر فترات عدم التوفر التي تريد حذفها:</p>
                    <div style="max-height: 300px; overflow-y: auto; margin-top: 15px;">
                        ${professeur.indisponibilites.map(([jour, periode], index) => `
                            <div class="checkbox-group" style="margin-bottom: 10px;">
                                <input type="checkbox" id="indispo_${index}" value="${index}">
                                <label for="indispo_${index}" style="color: black; font-weight: normal;">
                                    <i class="fas fa-calendar-day"></i>
                                    ${jour} - ${periode === 'matin' ? 'صباح' : 'مساء'}
                                </label>
                            </div>
                        `).join('')}
                    </div>
                </div>
            `,
            showCancelButton: true,
            confirmButtonText: 'حذف المحدد',
            cancelButtonText: 'إلغاء',
            confirmButtonColor: '#d33',
            preConfirm: () => {
                const checkboxes = document.querySelectorAll('input[type="checkbox"]:checked');
                return Array.from(checkboxes).map(cb => parseInt(cb.value));
            }
        }).then((result) => {
            if (result.isConfirmed && result.value.length > 0) {
                // Trier les indices en ordre décroissant pour éviter les problèmes d'index
                const indices = result.value.sort((a, b) => b - a);
                
                indices.forEach(index => {
                    professeur.indisponibilites.splice(index, 1);
                });
                
                this.afficherProfesseurs();
                this.updateStats();
                this.saveToLocalStorage();
                
                Swal.fire('نجاح', `تم حذف ${indices.length} فترة عدم توفر`, 'success');
            }
        });
    }

    importerExcel() {
        const input = document.createElement('input');
        input.type = 'file';
        input.accept = '.xlsx,.xls';
        
        input.onchange = (e) => {
            const file = e.target.files[0];
            if (!file) return;
            
            const reader = new FileReader();
            reader.onload = (e) => {
                try {
                    const data = new Uint8Array(e.target.result);
                    const workbook = XLSX.read(data, { type: 'array' });
                    const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
                    const jsonData = XLSX.utils.sheet_to_json(firstSheet, { header: 1 });
                    
                    if (jsonData.length === 0) {
                        Swal.fire('خطأ', 'الملف فارغ', 'error');
                        return;
                    }
                    
                    // Trouver les indices des colonnes
                    const headers = jsonData[0];
                    let nomIndex = -1;
                    let matiereIndex = -1;
                    
                    // Chercher les en-têtes
                    headers.forEach((header, index) => {
                        const headerStr = String(header || '').trim();
                        if (headerStr.includes('اسم') || headerStr.includes('الاسم') || headerStr.includes('name') || headerStr.includes('Name')) {
                            nomIndex = index;
                        }
                        if (headerStr.includes('مادة') || headerStr.includes('المادة') || headerStr.includes('subject') || headerStr.includes('Subject') || headerStr.includes('matière')) {
                            matiereIndex = index;
                        }
                    });
                    
                    // Si les en-têtes ne sont pas trouvés, utiliser les premières colonnes
                    if (nomIndex === -1 && headers.length > 0) nomIndex = 0;
                    if (matiereIndex === -1 && headers.length > 1) matiereIndex = 1;
                    
                    if (nomIndex === -1 || matiereIndex === -1) {
                        Swal.fire('خطأ', 
                            'الملف لا يحتوي على الأعمدة المطلوبة (الاسم الكامل، المادة)',
                            'error');
                        return;
                    }
                    
                    // Traiter les données (sans l'en-tête)
                    this.professeurs = [];
                    for (let i = 1; i < jsonData.length; i++) {
                        const row = jsonData[i];
                        if (!row || row.length === 0) continue;
                        
                        const nom = String(row[nomIndex] || '').trim();
                        const matiere = String(row[matiereIndex] || '').trim();
                        
                        if (nom && matiere) {
                            this.professeurs.push({
                                nom: nom,
                                matiere: matiere,
                                numero: '',
                                indisponibilites: []
                            });
                        }
                    }
                    
                    this.afficherProfesseurs();
                    this.updateStats();
                    this.saveToLocalStorage();
                    
                    // Extraire et afficher les matières
                    const matieres = this.obtenirListeMatieres();
                    let message = `تم استيراد ${this.professeurs.length} أستاذ بنجاح!`;
                    
                    if (matieres.length > 0) {
                        message += `\n\n📚 تم اكتشاف ${matieres.length} مادة:`;
                        matieres.slice(0, 5).forEach(([matiere, count]) => {
                            message += `\n• ${matiere} (${count} أستاذ)`;
                        });
                        if (matieres.length > 5) message += '\n...';
                    }
                    
                    Swal.fire('نجاح', message, 'success');
                    
                } catch (error) {
                    console.error('Error reading file:', error);
                    Swal.fire('خطأ', `خطأ في قراءة الملف: ${error.message}`, 'error');
                }
            };
            
            reader.readAsArrayBuffer(file);
        };
        
        input.click();
    }

    // ========== GESTION DES MATIÈRES ==========

    obtenirListeMatieres() {
        if (!this.professeurs.length) return [];
        
        const matieresCount = {};
        this.professeurs.forEach(prof => {
            const matiere = typeof prof.matiere === 'string' ? prof.matiere.trim() : String(prof.matiere || '').trim();
            if (matiere) {
                matieresCount[matiere] = (matieresCount[matiere] || 0) + 1;
            }
        });
        
        return Object.entries(matieresCount).sort((a, b) => b[1] - a[1]);
    }

    extraireMatieres() {
        if (!this.professeurs.length) {
            Swal.fire('تنبيه', 'لا يوجد أساتذة. يرجى استيراد الأساتذة أولاً.', 'warning');
            return;
        }
        
        const matieres = this.obtenirListeMatieres();
        const tbody = document.getElementById('extractedSubjectsBody');
        tbody.innerHTML = '';
        this.selectedSubjects.clear();
        
        matieres.forEach(([matiere, count], index) => {
            const tr = document.createElement('tr');
            tr.dataset.index = index;
            tr.dataset.matiere = matiere;
            tr.innerHTML = `
                <td><input type="checkbox" id="subject_${index}"></td>
                <td>${matiere}</td>
                <td>${count}</td>
            `;
            
            tr.addEventListener('click', (e) => {
                if (e.target.type !== 'checkbox') {
                    const checkbox = tr.querySelector('input[type="checkbox"]');
                    checkbox.checked = !checkbox.checked;
                    this.toggleSubjectSelection(checkbox.checked, matiere);
                }
            });
            
            const checkbox = tr.querySelector('input[type="checkbox"]');
            checkbox.addEventListener('change', (e) => {
                this.toggleSubjectSelection(e.target.checked, matiere);
            });
            
            tbody.appendChild(tr);
        });
        
        this.showModal('subjectsExtractModal');
    }

    toggleSubjectSelection(checked, matiere) {
        if (checked) {
            this.selectedSubjects.add(matiere);
        } else {
            this.selectedSubjects.delete(matiere);
        }
    }

    utiliserMatieresSelectionnees() {
        if (this.selectedSubjects.size === 0) {
            Swal.fire('تنبيه', 'يرجى اختيار مادة واحدة على الأقل', 'warning');
            return;
        }
        
        const matieresArray = Array.from(this.selectedSubjects);
        const nbrMatieres = matieresArray.length;
        
        document.getElementById('subjectCount').value = nbrMatieres;
        this.genererFormulaireMatieres();
        
        // Remplir le formulaire avec les matières sélectionnées
        const tbody = document.getElementById('subjectsTableBody');
        if (!tbody) return;
        
        for (let i = 0; i < nbrMatieres; i++) {
            if (i < this.matieres.length) {
                this.matieres[i].nom = matieresArray[i];
                this.matieres[i].date = `${(i % 30) + 1}/12/2024`;
            }
        }
        
        this.remplirFormulaireMatieres();
        this.closeModal('subjectsExtractModal');
        
        Swal.fire('نجاح', `تم تحميل ${nbrMatieres} مادة من قائمة الأساتذة`, 'success');
    }

    supprimerMatieresSelectionnees() {
        if (this.selectedSubjects.size === 0) {
            Swal.fire('تنبيه', 'يرجى اختيار مادة واحدة على الأقل للحذف', 'warning');
            return;
        }
        
        const matieresArray = Array.from(this.selectedSubjects);
        Swal.fire({
            title: 'تأكيد الحذف',
            html: `هل أنت متأكد من حذف ${matieresArray.length} مادة؟<br><br>${matieresArray.map(m => `• ${m}`).join('<br>')}`,
            icon: 'warning',
            showCancelButton: true,
            confirmButtonText: 'نعم، احذف',
            cancelButtonText: 'إلغاء',
            confirmButtonColor: '#d33'
        }).then((result) => {
            if (result.isConfirmed) {
                let count = 0;
                this.professeurs.forEach(prof => {
                    const profMatiere = typeof prof.matiere === 'string' ? prof.matiere.trim() : String(prof.matiere || '').trim();
                    if (matieresArray.includes(profMatiere)) {
                        prof.matiere = '';
                        count++;
                    }
                });
                
                this.afficherProfesseurs();
                this.saveToLocalStorage();
                this.closeModal('subjectsExtractModal');
                
                Swal.fire('نجاح', `تم حذف ${count} مادة بنجاح`, 'success');
            }
        });
    }

    remplirMatieresDepuisProfesseurs() {
        const matieres = this.obtenirListeMatieres();
        if (matieres.length === 0) {
            Swal.fire('تنبيه', 'لا توجد مواد في قائمة الأساتذة', 'warning');
            return;
        }
        
        this.selectedSubjects.clear();
        matieres.forEach(([matiere]) => this.selectedSubjects.add(matiere));
        this.utiliserMatieresSelectionnees();
    }

    genererFormulaireMatieres() {
        try {
            const nbrMatieres = parseInt(document.getElementById('subjectCount').value) || 0;
            if (nbrMatieres < 0) {
                Swal.fire('خطأ', 'عدد المواد يجب أن يكون موجباً أو صفراً', 'error');
                return;
            }
            
            const container = document.getElementById('subjectsFormContainer');
            container.innerHTML = '';
            
            if (nbrMatieres === 0) {
                // Si 0 matières, vider le tableau
                this.matieres = [];
                this.saveToLocalStorage();
                return;
            }
            
            // Créer le tableau
            const table = document.createElement('table');
            table.className = 'data-table';
            table.innerHTML = `
                <thead>
                    <tr>
                        <th>اسم المادة</th>
                        <th>تاريخ الامتحان</th>
                        <th>الفترة</th>
                        <th>وقت البدء</th>
                        <th>وقت الانتهاء</th>
                        <th>المدة</th>
                    </tr>
                </thead>
                <tbody id="subjectsTableBody"></tbody>
            `;
            
            container.appendChild(table);
            
            // Initialiser ou ajuster le tableau des matières
            if (this.matieres.length > nbrMatieres) {
                this.matieres = this.matieres.slice(0, nbrMatieres);
            } else {
                while (this.matieres.length < nbrMatieres) {
                    const i = this.matieres.length;
                    const dateExam = `${(i % 30) + 1}/12/2024`;
                    const periode = i % 2 === 0 ? 'صباح' : 'مساء';
                    const heureDebut = i % 2 === 0 ? '08:00' : '14:00';
                    const heureFin = i % 2 === 0 ? '12:00' : '18:00';
                    const duree = i % 2 === 0 ? '04:00' : '04:00';
                    
                    this.matieres.push({
                        nom: '',
                        date: dateExam,
                        periode: periode,
                        heure_debut: heureDebut,
                        heure_fin: heureFin,
                        duree: duree
                    });
                }
            }
            
            const tbody = document.getElementById('subjectsTableBody');
            tbody.innerHTML = '';
            
            this.matieres.forEach((matiere, i) => {
                const tr = document.createElement('tr');
                tr.innerHTML = `
                    <td><input type="text" class="form-control" data-field="nom" data-index="${i}" value="${matiere.nom || ''}"></td>
                    <td><input type="text" class="form-control" data-field="date" data-index="${i}" value="${matiere.date}"></td>
                    <td>
                        <select class="form-control" data-field="periode" data-index="${i}">
                            <option value="صباح" ${matiere.periode === 'صباح' ? 'selected' : ''}>صباح</option>
                            <option value="مساء" ${matiere.periode === 'مساء' ? 'selected' : ''}>مساء</option>
                        </select>
                    </td>
                    <td><input type="time" class="form-control" data-field="heure_debut" data-index="${i}" value="${matiere.heure_debut}"></td>
                    <td><input type="time" class="form-control" data-field="heure_fin" data-index="${i}" value="${matiere.heure_fin}"></td>
                    <td><input type="text" class="form-control" data-field="duree" data-index="${i}" value="${matiere.duree}" readonly></td>
                `;
                
                tbody.appendChild(tr);
            });
            
            // Ajouter les événements pour le calcul de la durée
            tbody.addEventListener('change', (e) => {
                const target = e.target;
                const index = parseInt(target.dataset.index);
                
                if (!isNaN(index) && index >= 0 && index < this.matieres.length) {
                    if (target.dataset.field === 'periode') {
                        this.updateHeuresPeriode(index, target.value);
                    } else if (target.dataset.field === 'heure_debut' || target.dataset.field === 'heure_fin') {
                        this.calculerDuree(index);
                    }
                    
                    this.sauvegarderMatiere(index);
                }
            });
            
            tbody.addEventListener('input', (e) => {
                const target = e.target;
                const index = parseInt(target.dataset.index);
                
                if (!isNaN(index) && index >= 0 && index < this.matieres.length && target.type === 'text') {
                    this.sauvegarderMatiere(index);
                }
            });
            
            this.saveToLocalStorage();
            
        } catch (error) {
            Swal.fire('خطأ', `خطأ في توليد نموذج المواد: ${error.message}`, 'error');
        }
    }

    updateHeuresPeriode(index, periode) {
        const heureDebutInput = document.querySelector(`[data-field="heure_debut"][data-index="${index}"]`);
        const heureFinInput = document.querySelector(`[data-field="heure_fin"][data-index="${index}"]`);
        const dureeInput = document.querySelector(`[data-field="duree"][data-index="${index}"]`);
        
        if (periode === 'صباح') {
            heureDebutInput.value = '08:00';
            heureFinInput.value = '12:00';
        } else {
            heureDebutInput.value = '14:00';
            heureFinInput.value = '18:00';
        }
        
        dureeInput.value = '04:00';
        
        if (index < this.matieres.length) {
            this.matieres[index].periode = periode;
            this.matieres[index].heure_debut = heureDebutInput.value;
            this.matieres[index].heure_fin = heureFinInput.value;
            this.matieres[index].duree = '04:00';
        }
        
        this.saveToLocalStorage();
    }

    calculerDuree(index) {
        const heureDebutInput = document.querySelector(`[data-field="heure_debut"][data-index="${index}"]`);
        const heureFinInput = document.querySelector(`[data-field="heure_fin"][data-index="${index}"]`);
        const dureeInput = document.querySelector(`[data-field="duree"][data-index="${index}"]`);
        
        if (!heureDebutInput || !heureFinInput || !dureeInput) return;
        
        if (!heureDebutInput.value || !heureFinInput.value) {
            dureeInput.value = '00:00';
            return;
        }
        
        const [h1, m1] = heureDebutInput.value.split(':').map(Number);
        const [h2, m2] = heureFinInput.value.split(':').map(Number);
        
        let totalMinutes = (h2 * 60 + m2) - (h1 * 60 + m1);
        if (totalMinutes < 0) totalMinutes += 24 * 60;
        
        const heures = Math.floor(totalMinutes / 60);
        const minutes = totalMinutes % 60;
        
        dureeInput.value = `${heures.toString().padStart(2, '0')}:${minutes.toString().padStart(2, '0')}`;
        
        if (index < this.matieres.length) {
            this.matieres[index].duree = dureeInput.value;
        }
        
        this.saveToLocalStorage();
    }

    sauvegarderMatiere(index) {
        if (index >= this.matieres.length) return;
        
        const inputs = document.querySelectorAll(`[data-index="${index}"]`);
        inputs.forEach(input => {
            const field = input.dataset.field;
            this.matieres[index][field] = input.value;
        });
        
        this.saveToLocalStorage();
    }

    remplirFormulaireMatieres() {
        const tbody = document.getElementById('subjectsTableBody');
        if (!tbody) return;
        
        this.matieres.forEach((matiere, index) => {
            const inputs = document.querySelectorAll(`[data-index="${index}"]`);
            inputs.forEach(input => {
                const field = input.dataset.field;
                input.value = matiere[field] || '';
            });
        });
    }

    supprimerMatiere() {
        if (this.matieres.length === 0) {
            Swal.fire('تنبيه', 'لا توجد مواد للحذف', 'warning');
            return;
        }
        
        Swal.fire({
            title: 'حذف مادة',
            input: 'select',
            inputOptions: this.matieres.reduce((options, matiere, index) => {
                options[index] = matiere.nom || `مادة ${index + 1}`;
                return options;
            }, {}),
            inputPlaceholder: 'اختر المادة للحذف',
            showCancelButton: true,
            confirmButtonText: 'حذف',
            cancelButtonText: 'إلغاء',
            confirmButtonColor: '#d33'
        }).then((result) => {
            if (result.isConfirmed) {
                const index = parseInt(result.value);
                this.matieres.splice(index, 1);
                document.getElementById('subjectCount').value = this.matieres.length;
                this.genererFormulaireMatieres();
                this.remplirFormulaireMatieres();
                this.saveToLocalStorage();
                
                Swal.fire('نجاح', 'تم حذف المادة بنجاح', 'success');
            }
        });
    }

    // ========== GÉNÉRATION DES AFFECTATIONS ==========

    genererAffectations() {
        try {
            if (this.professeurs.length === 0) {
                Swal.fire('خطأ', 'يرجى استيراد أو إضافة أساتذة أولاً', 'error');
                return;
            }
            
            const nbrSalles = parseInt(document.getElementById('roomCount').value) || 0;
            const nbrProfsSalle = parseInt(document.getElementById('profsPerRoom').value) || 2;
            
            // Filtrer les matières avec nom
            const matieresValides = this.matieres.filter(m => m.nom && m.nom.trim());
            
            if (matieresValides.length === 0) {
                Swal.fire('خطأ', 'يرجى ملء المواد أولاً', 'error');
                return;
            }
            
            this.type_examen = document.getElementById('examType').value.trim() || "امتحان عادي";
            
            if (nbrSalles <= 0) {
                Swal.fire('خطأ', 'يرجى إدخال عدد القاعات', 'error');
                return;
            }
            
            this.salles = Array.from({ length: nbrSalles }, (_, i) => `القاعة ${i + 1}`);
            this.affectations = this.calculerAffectations(nbrProfsSalle, matieresValides);
            
            this.afficherResultats();
            this.afficherStatsAffectations();
            this.afficherProfsNonAffectes();
            this.saveToLocalStorage();
            
            Swal.fire('نجاح', 'تم توليد التوزيع بنجاح', 'success');
            
        } catch (error) {
            Swal.fire('خطأ', `خطأ في التوليد: ${error.message}`, 'error');
        }
    }

    calculerAffectations(nbrProfsSalle, matieres) {
        const affectations = [];
        const historiqueProfs = {};
        
        matieres.forEach(matiere => {
            const demiJournee = `${matiere.date}-${matiere.periode}`;
            const dateHeure = `${matiere.date} ${matiere.heure_debut}-${matiere.heure_fin}`;
            
            this.salles.forEach(salle => {
                const profsAffectes = [];
                const profsDisponibles = [...this.professeurs];
                
                // Mélanger les professeurs
                for (let i = profsDisponibles.length - 1; i > 0; i--) {
                    const j = Math.floor(Math.random() * (i + 1));
                    [profsDisponibles[i], profsDisponibles[j]] = [profsDisponibles[j], profsDisponibles[i]];
                }
                
                for (let i = 0; i < nbrProfsSalle; i++) {
                    let profTrouve = null;
                    
                    for (const prof of profsDisponibles) {
                        if (!profsAffectes.includes(prof) && 
                            this.estProfesseurValide(prof, salle, demiJournee, matiere.nom, affectations, historiqueProfs)) {
                            profTrouve = prof;
                            break;
                        }
                    }
                    
                    if (profTrouve) {
                        profsAffectes.push(profTrouve);
                        
                        // Utiliser le nom comme identifiant au lieu du numéro de location
                        const profId = profTrouve.nom;
                        
                        if (!historiqueProfs[profId]) {
                            historiqueProfs[profId] = {
                                salles: new Set(),
                                matieres: new Set(),
                                creneaux: new Set()
                            };
                        }
                        
                        historiqueProfs[profId].salles.add(salle);
                        historiqueProfs[profId].matieres.add(matiere.nom);
                        historiqueProfs[profId].creneaux.add(demiJournee);
                        
                        affectations.push({
                            matiere: matiere.nom,
                            date_heure: dateHeure,
                            salle: salle,
                            professeur: profTrouve.nom,
                            periode: matiere.periode
                        });
                    }
                }
            });
        });
        
        return affectations;
    }

    estProfesseurValide(prof, salle, demiJournee, matiere, affectationsExistantes, historique) {
        const [date, periode] = demiJournee.split('-');
        
        // Vérifier les indisponibilités
        const indispos = prof.indisponibilites || [];
        if (indispos.some(([jour, p]) => jour === date && p === periode)) {
            return false;
        }
        
        // Vérifier les contraintes
        const pasPropreMatiere = document.getElementById('constraint1').checked;
        const pasMemeSalle = document.getElementById('constraint2').checked;
        const pasMemeGroupe = document.getElementById('constraint3').checked;
        
        // Utiliser le nom comme identifiant
        const profId = prof.nom;
        
        // Convertir la matière du professeur en chaîne de caractères
        const profMatiere = typeof prof.matiere === 'string' ? prof.matiere.trim() : String(prof.matiere || '').trim();
        
        if (pasPropreMatiere && profMatiere === matiere) {
            return false;
        }
        
        if (pasMemeSalle && historique[profId] && historique[profId].salles.has(salle)) {
            return false;
        }
        
        if (pasMemeGroupe && historique[profId] && historique[profId].creneaux.has(demiJournee)) {
            return false;
        }
        
        return true;
    }

    afficherResultats() {
        const tbody = document.getElementById('assignmentsTableBody');
        tbody.innerHTML = '';
        
        if (this.affectations.length === 0) return;
        
        const nbrProfsSalle = parseInt(document.getElementById('profsPerRoom').value) || 2;
        const affectationsParMatiere = {};
        
        // Organiser les affectations
        this.affectations.forEach(affectation => {
            const matiere = affectation.matiere;
            if (!affectationsParMatiere[matiere]) {
                affectationsParMatiere[matiere] = {};
            }
            
            const salle = affectation.salle;
            if (!affectationsParMatiere[matiere][salle]) {
                affectationsParMatiere[matiere][salle] = {
                    date_heure: affectation.date_heure,
                    professeurs: []
                };
            }
            
            affectationsParMatiere[matiere][salle].professeurs.push(affectation.professeur);
        });
        
        // Mettre à jour l'en-tête du tableau (de droite à gauche)
        const thead = document.querySelector('#assignmentsTable thead tr');
        thead.innerHTML = '<th>الحالة</th>';
        
        // Ajouter les colonnes pour chaque professeur (de droite à gauche)
        for (let i = nbrProfsSalle; i >= 1; i--) {
            thead.innerHTML += `<th>الأستاذ ${i}</th>`;
        }
        
        thead.innerHTML += '<th>القاعة</th><th>تاريخ ووقت الامتحان</th><th>المادة</th>';
        
        // Afficher les données
        for (const [matiere, salles] of Object.entries(affectationsParMatiere)) {
            for (const [salle, info] of Object.entries(salles)) {
                const tr = document.createElement('tr');
                let rowHTML = '';
                
                // Ajouter la colonne état (première colonne à droite)
                const statut = info.professeurs.length >= nbrProfsSalle ? '🟢 مكتمل' : '🟡 جزئي';
                rowHTML += `<td>${statut}</td>`;
                
                // Ajouter les professeurs dans des colonnes séparées (de droite à gauche)
                for (let i = nbrProfsSalle - 1; i >= 0; i--) {
                    rowHTML += `<td>${info.professeurs[i] || ''}</td>`;
                }
                
                // Ajouter les autres colonnes
                rowHTML += `<td>${salle}</td><td>${info.date_heure}</td><td>${matiere}</td>`;
                
                tr.innerHTML = rowHTML;
                tbody.appendChild(tr);
            }
        }
    }

    afficherStatsAffectations() {
        const container = document.getElementById('assignmentsStats');
        container.innerHTML = '';
        
        if (this.affectations.length === 0) return;
        
        const totalAffectations = this.affectations.length;
        const matieresUniques = new Set(this.affectations.map(a => a.matiere)).size;
        const profsAffectes = new Set(this.affectations.map(a => a.professeur)).size;
        const sallesUtilisees = new Set(this.affectations.map(a => a.salle)).size;
        const totalProfs = this.professeurs.length;
        const profsNonAffectes = totalProfs - profsAffectes;
        
        const stats = [
            { icon: 'fas fa-list', title: 'إجمالي التوزيعات', value: totalAffectations },
            { icon: 'fas fa-book', title: 'المواد', value: matieresUniques },
            { icon: 'fas fa-users', title: 'الأساتذة الموزعين', value: profsAffectes },
            { icon: 'fas fa-user-times', title: 'الأساتذة غير الموزعين', value: profsNonAffectes },
            { icon: 'fas fa-school', title: 'القاعات المستخدمة', value: sallesUtilisees }
        ];
        
        stats.forEach(stat => {
            const card = document.createElement('div');
            card.className = 'stat-card';
            card.innerHTML = `
                <div class="stat-icon">
                    <i class="${stat.icon}"></i>
                </div>
                <h3>${stat.title}</h3>
                <p>${stat.value}</p>
            `;
            container.appendChild(card);
        });
    }

    // ========== NOUVELLE MÉTHODE : AFFICHER LES PROFESSEURS NON AFFECTÉS ==========

    afficherProfsNonAffectes() {
        if (this.affectations.length === 0) return;
        
        // Trouver ou créer la section pour afficher les non-affectés
        let nonAffectesSection = document.getElementById('nonAffectesSection');
        if (!nonAffectesSection) {
            nonAffectesSection = document.createElement('div');
            nonAffectesSection.id = 'nonAffectesSection';
            nonAffectesSection.className = 'card';
            nonAffectesSection.innerHTML = `
                <div class="card-header">
                    <h2><i class="fas fa-user-times"></i> الأساتذة غير الموزعين حسب المادة</h2>
                </div>
                <div class="table-container" id="nonAffectesContainer"></div>
            `;
            
            // Insérer après le tableau des affectations
            const assignmentsTable = document.querySelector('#affectations .table-container');
            assignmentsTable.parentNode.insertBefore(nonAffectesSection, assignmentsTable.nextSibling);
        }
        
        const container = document.getElementById('nonAffectesContainer');
        container.innerHTML = '';
        
        // Obtenir la liste de tous les professeurs affectés
        const profsAffectes = new Set(this.affectations.map(a => a.professeur));
        
        // Grouper les professeurs non affectés par matière
        const matieresNonAffectes = {};
        const profsSansMatiere = [];
        
        this.professeurs.forEach(prof => {
            if (!profsAffectes.has(prof.nom)) {
                const matiere = prof.matiere ? prof.matiere.trim() : 'غير محدد';
                if (matiere && matiere !== 'غير محدد') {
                    if (!matieresNonAffectes[matiere]) {
                        matieresNonAffectes[matiere] = [];
                    }
                    matieresNonAffectes[matiere].push(prof);
                } else {
                    profsSansMatiere.push(prof);
                }
            }
        });
        
        // Créer le tableau principal
        const table = document.createElement('table');
        table.className = 'data-table';
        table.innerHTML = `
            <thead>
                <tr>
                    <th>المادة</th>
                    <th>الاسم الكامل</th>
                    <th>رقم التأجير</th>
                    <th>فترات عدم التوفر</th>
                    <th>الحالة</th>
                </tr>
            </thead>
            <tbody id="nonAffectesTableBody"></tbody>
        `;
        
        container.appendChild(table);
        const tbody = document.getElementById('nonAffectesTableBody');
        
        // Ajouter les professeurs non affectés par matière
        Object.entries(matieresNonAffectes).forEach(([matiere, profs]) => {
            // Ajouter une ligne pour le nom de la matière
            const matiereRow = document.createElement('tr');
            matiereRow.className = 'matiere-header-row';
            matiereRow.innerHTML = `
                <td colspan="5" style="background-color: #2a2a3c; font-weight: bold; color: #339af0;">
                    <i class="fas fa-book"></i> ${matiere} (${profs.length} أستاذ)
                </td>
            `;
            tbody.appendChild(matiereRow);
            
            // Ajouter les professeurs de cette matière
            profs.forEach((prof, index) => {
                const indispoText = prof.indisponibilites && prof.indisponibilites.length > 0 
                    ? prof.indisponibilites.map(([jour, periode]) => `${jour}-${periode === 'matin' ? 'صباح' : 'مساء'}`).join(', ')
                    : 'لا يوجد';
                
                const tr = document.createElement('tr');
                tr.className = 'non-affecte-row';
                tr.innerHTML = `
                    <td></td>
                    <td>${prof.nom}</td>
                    <td>${prof.numero || ''}</td>
                    <td>${indispoText}</td>
                    <td><span class="status-badge status-unavailable">غير موزع</span></td>
                `;
                tbody.appendChild(tr);
            });
            
            // Ajouter une ligne vide entre les matières
            const emptyRow = document.createElement('tr');
            emptyRow.innerHTML = '<td colspan="5" style="height: 10px;"></td>';
            tbody.appendChild(emptyRow);
        });
        
        // Ajouter les professeurs sans matière spécifiée
        if (profsSansMatiere.length > 0) {
            const sansMatiereRow = document.createElement('tr');
            sansMatiereRow.className = 'matiere-header-row';
            sansMatiereRow.innerHTML = `
                <td colspan="5" style="background-color: #2a2a3c; font-weight: bold; color: #fa5252;">
                    <i class="fas fa-question-circle"></i> أساتذة بدون مادة محددة (${profsSansMatiere.length} أستاذ)
                </td>
            `;
            tbody.appendChild(sansMatiereRow);
            
            profsSansMatiere.forEach(prof => {
                const indispoText = prof.indisponibilites && prof.indisponibilites.length > 0 
                    ? prof.indisponibilites.map(([jour, periode]) => `${jour}-${periode === 'matin' ? 'صباح' : 'مساء'}`).join(', ')
                    : 'لا يوجد';
                
                const tr = document.createElement('tr');
                tr.className = 'non-affecte-row';
                tr.innerHTML = `
                    <td>غير محدد</td>
                    <td>${prof.nom}</td>
                    <td>${prof.numero || ''}</td>
                    <td>${indispoText}</td>
                    <td><span class="status-badge status-unavailable">غير موزع</span></td>
                `;
                tbody.appendChild(tr);
            });
        }
        
        // Ajouter un résumé
        const totalNonAffectes = this.professeurs.length - profsAffectes.size;
        const summaryRow = document.createElement('tr');
        summaryRow.className = 'summary-row';
        summaryRow.innerHTML = `
            <td colspan="5" style="background-color: #495057; color: white; font-weight: bold; text-align: center;">
                <i class="fas fa-chart-bar"></i> إجمالي الأساتذة غير الموزعين: ${totalNonAffectes} من ${this.professeurs.length}
            </td>
        `;
        tbody.appendChild(summaryRow);
    }

    // ========== EXPORT EXCEL ==========

    showExcelConfigModal() {
        if (this.affectations.length === 0) {
            Swal.fire('تنبيه', 'يرجى توليد التوزيع أولاً', 'warning');
            return;
        }
        
        document.getElementById('universityName').value = '';
        document.getElementById('facultyName').value = '';
        document.getElementById('faculty').value = '';
        document.getElementById('academicYear').value = '';
        document.getElementById('ecoleName').value = '';
        
        this.showModal('excelConfigModal');
    }

    genererExcel() {
        const university = document.getElementById('universityName').value.trim();
        const faculty = document.getElementById('facultyName').value.trim();
        const facultyy = document.getElementById('faculty').value.trim();
        const academicYear = document.getElementById('academicYear').value.trim();
        const ecole = document.getElementById('ecoleName').value.trim();

        if (!university || !faculty || !academicYear || !ecole) {
            Swal.fire('خطأ', 'يرجى ملء جميع الحقول الإلزامية', 'error');
            return;
        }

        try {
            const wb = XLSX.utils.book_new();
            const nbrProfsSalle = parseInt(document.getElementById('profsPerRoom').value) || 2;

            /* =========================
               FEUILLE PRINCIPALE : الجدول العام
            ========================= */
            const data = [];
            data.push([`الوزارة : ${facultyy}`]);
            data.push([`المديرية: ${university}`]);
            data.push([`الأكادمية: ${faculty}`]);
            data.push([`المؤسسة: ${ecole}`]);
            data.push([ `السنة الدراسية : ${academicYear}`]);
            data.push([`نوع الامتحان : ${this.type_examen}`]);
            data.push([]);
            
            const header = ['المادة', 'تاريخ ووقت الامتحان', 'القاعة'];
            for (let i = nbrProfsSalle; i >= 1; i--) {
                header.push(`الأستاذ ${i}`);
            }
            data.push(header);

            // Grouper les affectations
            const grouped = {};
            this.affectations.forEach(a => {
                const key = `${a.matiere}_${a.salle}`;
                if (!grouped[key]) {
                    grouped[key] = {
                        matiere: a.matiere,
                        salle: a.salle,
                        date: a.date_heure,
                        profs: []
                    };
                }
                grouped[key].profs.push(a.professeur);
            });

            Object.values(grouped).forEach(item => {
                const row = [item.matiere, item.date, item.salle];
                for (let i = nbrProfsSalle - 1; i >= 0; i--) {
                    row.push(item.profs[i] || '');
                }
                data.push(row);
            });

            const wsGlobal = XLSX.utils.aoa_to_sheet(data);
            wsGlobal['!rtl'] = true;
            const range = XLSX.utils.decode_range(wsGlobal['!ref']);

            for (let R = range.s.r; R <= range.e.r; R++) {
                for (let C = range.s.c; C <= range.e.c; C++) {
                    const ref = XLSX.utils.encode_cell({ r: R, c: C });
                    if (!wsGlobal[ref]) continue;

                    wsGlobal[ref].s = {
                        alignment: { horizontal: 'center', vertical: 'center', readingOrder: 'rtl' },
                        font: { name: 'Arial', sz: R <= 6 ? 14 : 11, bold: R === 7 },
                        fill: R === 7 ? { fgColor: { rgb: "D9E1F2" } } : 
                              R <= 6 ? { fgColor: { rgb: "BDD7EE" } } : undefined,
                        border: {
                            top: { style: "thin", color: { rgb: "000000" } },
                            bottom: { style: "thin", color: { rgb: "000000" } },
                            left: { style: "thin", color: { rgb: "000000" } },
                            right: { style: "thin", color: { rgb: "000000" } }
                        }
                    };
                }
            }

            const colWidths = [];
            for (let i = 0; i < nbrProfsSalle; i++) colWidths.push({ wch: 25 });
            colWidths.push({ wch: 15 }, { wch: 25 }, { wch: 25 });
            wsGlobal['!cols'] = colWidths;

            XLSX.utils.book_append_sheet(wb, wsGlobal, 'الجدول العام');

            /* =========================
               FEUILLE : الأساتذة غير الموزعين
            ========================= */
            const nonAffectesData = [];
            nonAffectesData.push([`الوزارة : ${facultyy}`]);
            nonAffectesData.push([`المديرية: ${university}`]);
            nonAffectesData.push([`الأكادمية: ${faculty}`]);
            nonAffectesData.push([`المؤسسة: ${ecole}`]);
            nonAffectesData.push([ `السنة الدراسية : ${academicYear}`]);
            nonAffectesData.push([`نوع الامتحان : ${this.type_examen}`]);
            nonAffectesData.push([]);
            nonAffectesData.push(['قائمة الأساتذة غير الموزعين لكل مادة']);
            nonAffectesData.push([]);

            // Obtenir la liste de tous les professeurs affectés
            const profsAffectes = new Set(this.affectations.map(a => a.professeur));
            
            // Grouper les professeurs non affectés par matière
            const matieresNonAffectes = {};
            
            this.professeurs.forEach(prof => {
                if (!profsAffectes.has(prof.nom)) {
                    const matiere = prof.matiere || 'غير محدد';
                    if (!matieresNonAffectes[matiere]) {
                        matieresNonAffectes[matiere] = [];
                    }
                    matieresNonAffectes[matiere].push(prof.nom);
                }
            });

            // Ajouter les matières avec professeurs non affectés
            Object.entries(matieresNonAffectes).forEach(([matiere, profs]) => {
                nonAffectesData.push([`المادة: ${matiere}`]);
                nonAffectesData.push(['الاسم الكامل', 'رقم التأجير', 'ملاحظات']);
                
                profs.forEach(profNom => {
                    const prof = this.professeurs.find(p => p.nom === profNom);
                    nonAffectesData.push([
                        profNom,
                        prof.numero || '',
                        prof.indisponibilites && prof.indisponibilites.length > 0 ? 'لديه عدم توفر' : ''
                    ]);
                });
                
                nonAffectesData.push([]);
            });

            // Ajouter les statistiques
            nonAffectesData.push([]);
            nonAffectesData.push(['إحصائيات عامة:']);
            nonAffectesData.push([`إجمالي الأساتذة: ${this.professeurs.length}`]);
            nonAffectesData.push([`عدد الأساتذة الموزعين: ${profsAffectes.size}`]);
            nonAffectesData.push([`عدد الأساتذة غير الموزعين: ${this.professeurs.length - profsAffectes.size}`]);

            const wsNonAffectes = XLSX.utils.aoa_to_sheet(nonAffectesData);
            wsNonAffectes['!rtl'] = true;
            const rangeNA = XLSX.utils.decode_range(wsNonAffectes['!ref']);

            for (let R = rangeNA.s.r; R <= rangeNA.e.r; R++) {
                for (let C = rangeNA.s.c; C <= rangeNA.e.c; C++) {
                    const ref = XLSX.utils.encode_cell({ r: R, c: C });
                    if (!wsNonAffectes[ref]) continue;

                    wsNonAffectes[ref].s = {
                        alignment: { horizontal: 'center', vertical: 'center', readingOrder: 'rtl' },
                        font: { 
                            name: 'Arial', 
                            sz: R <= 6 ? 14 : (R === 7 ? 16 : 11),
                            bold: R === 7 || R === 8
                        },
                        fill: R === 7 ? { fgColor: { rgb: "FCE4D6" } } : 
                              R <= 6 ? { fgColor: { rgb: "E2EFDA" } } : undefined,
                        border: {
                            top: { style: "thin", color: { rgb: "000000" } },
                            bottom: { style: "thin", color: { rgb: "000000" } },
                            left: { style: "thin", color: { rgb: "000000" } },
                            right: { style: "thin", color: { rgb: "000000" } }
                        }
                    };
                }
            }

            wsNonAffectes['!cols'] = [{ wch: 30 }, { wch: 15 }, { wch: 20 }];
            XLSX.utils.book_append_sheet(wb, wsNonAffectes, 'غير الموزعين');

            /* =========================
               FEUILLES PAR MATIÈRE
            ========================= */
            const matieresGroup = {};
            this.affectations.forEach(a => {
                if (!matieresGroup[a.matiere]) matieresGroup[a.matiere] = [];
                matieresGroup[a.matiere].push(a);
            });

            for (const [matiere, affectationsMatiere] of Object.entries(matieresGroup)) {
                const dataMatiere = [];
                
                dataMatiere.push([`المديرية : ${university}`]);
                dataMatiere.push([`الأكادمية: ${faculty}`]);
                dataMatiere.push([`الوزارة: ${facultyy}`]);
                dataMatiere.push([`المؤسسة : ${ecole}`]);
                dataMatiere.push([`السنة الدراسية : ${academicYear}`]);
                dataMatiere.push([`نوع الامتحان : ${this.type_examen}`]);
                dataMatiere.push([`المادة : ${matiere}`]);
                dataMatiere.push([]);

                const headerMatiere = ['القاعة', 'تاريخ ووقت الامتحان'];
                for (let i = nbrProfsSalle; i >= 1; i--) headerMatiere.push(`الأستاذ ${i}`);
                dataMatiere.push(headerMatiere);

                const groupedBySalle = {};
                affectationsMatiere.forEach(a => {
                    if (!groupedBySalle[a.salle]) {
                        groupedBySalle[a.salle] = { salle: a.salle, date: a.date_heure, profs: [] };
                    }
                    groupedBySalle[a.salle].profs.push(a.professeur);
                });

                Object.values(groupedBySalle).forEach(item => {
                    const row = [item.salle, item.date];
                    for (let i = nbrProfsSalle - 1; i >= 0; i--) row.push(item.profs[i] || '');
                    dataMatiere.push(row);
                });

                const wsMatiere = XLSX.utils.aoa_to_sheet(dataMatiere);
                wsMatiere['!rtl'] = true;
                const rangeMatiere = XLSX.utils.decode_range(wsMatiere['!ref']);

                for (let R = rangeMatiere.s.r; R <= rangeMatiere.e.r; R++) {
                    for (let C = rangeMatiere.s.c; C <= rangeMatiere.e.c; C++) {
                        const ref = XLSX.utils.encode_cell({ r: R, c: C });
                        if (!wsMatiere[ref]) continue;

                        wsMatiere[ref].s = {
                            alignment: { horizontal: 'center', vertical: 'center', readingOrder: 'rtl' },
                            font: { name: 'Arial', sz: R <= 6 ? 14 : 11, bold: R === 7 },
                            fill: R === 7 ? { fgColor: { rgb: "E2EFDA" } } : 
                                  R <= 6 ? { fgColor: { rgb: "FCE4D6" } } : undefined,
                            border: {
                                top: { style: "thin", color: { rgb: "000000" } },
                                bottom: { style: "thin", color: { rgb: "000000" } },
                                left: { style: "thin", color: { rgb: "000000" } },
                                right: { style: "thin", color: { rgb: "000000" } }
                            }
                        };
                    }
                }

                const colWidthsMatiere = [];
                for (let i = 0; i < nbrProfsSalle; i++) colWidthsMatiere.push({ wch: 25 });
                colWidthsMatiere.push({ wch: 25 }, { wch: 15 });
                wsMatiere['!cols'] = colWidthsMatiere;

                const sheetName = matiere.substring(0, 31);
                XLSX.utils.book_append_sheet(wb, wsMatiere, sheetName);
            }

            /* =========================
               SAUVEGARDE
            ========================= */
            const fileName = `توزيع_الأساتذة_${new Date().toISOString().slice(0,10)}.xlsx`;
            XLSX.writeFile(wb, fileName);

            this.closeModal('excelConfigModal');
            Swal.fire('نجاح', `تم إنشاء ملف Excel بنجاح يحتوي على:<br>1. الجدول العام<br>2. قائمة غير الموزعين<br>3. جداول المواد`, 'success');

        } catch (e) {
            Swal.fire('خطأ', `خطأ في تصدير الإكسل: ${e.message}`, 'error');
        }
    }

    // ========== STATISTIQUES ==========

    updateStats() {
        const totalProfs = this.professeurs.length;
        const availableProfs = this.professeurs.filter(p => !p.indisponibilites || p.indisponibilites.length === 0).length;
        const unavailableProfs = totalProfs - availableProfs;
        
        document.getElementById('totalProfs').textContent = totalProfs;
        document.getElementById('availableProfs').textContent = availableProfs;
        document.getElementById('unavailableProfs').textContent = unavailableProfs;
        document.getElementById('profCount').textContent = `الأساتذة: ${totalProfs}`;
    }
}

// Initialiser l'application lorsque la page est chargée
document.addEventListener('DOMContentLoaded', () => {
    window.app = new ApplicationAffectation();
});