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
        document.getElementById('deleteAllProfessorsBtn').addEventListener('click', () => this.supprimerTousLesProfesseurs());
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
        document.querySelectorAll('.tab-content').forEach(tab => {
            tab.classList.remove('active');
        });
        document.querySelectorAll('.nav-tab').forEach(tab => {
            tab.classList.remove('active');
        });
        document.getElementById(tabName).classList.add('active');
        document.querySelector(`[data-tab="${tabName}"]`).classList.add('active');
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
        const statusElement = document.getElementById('statusMessage');
        if (statusElement) {
            statusElement.textContent = message;
        } else {
            console.log("Status: " + message);
        }
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
        document.querySelectorAll('#professorsTableBody tr').forEach(r => {
            r.classList.remove('selected-row');
        });
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
            professeur.indisponibilites = this.professeurs[this.selectedProfessorIndex].indisponibilites || [];
            this.professeurs[this.selectedProfessorIndex] = professeur;
        } else {
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

    supprimerTousLesProfesseurs() {
        if (this.professeurs.length === 0) {
            Swal.fire('تنبيه', 'لا يوجد أساتذة في القائمة', 'info');
            return;
        }
        
        Swal.fire({
            title: 'تأكيد الحذف الكامل',
            html: `
                <div style="text-align: right; margin: 20px 0;">
                    <p>هل أنت متأكد من حذف جميع الأساتذة؟</p>
                    <p style="color: #d33; font-weight: bold; margin-top: 10px;">
                        ⚠️ سيتم حذف ${this.professeurs.length} أستاذ من القائمة!
                    </p>
                    <p style="margin-top: 15px; color: #666;">
                        <i class="fas fa-info-circle"></i>
                        هذه العملية لا يمكن التراجع عنها.
                    </p>
                </div>
            `,
            icon: 'warning',
            showCancelButton: true,
            confirmButtonText: 'نعم، احذف الكل',
            cancelButtonText: 'إلغاء',
            confirmButtonColor: '#d33',
            focusCancel: true
        }).then((result) => {
            if (result.isConfirmed) {
                const count = this.professeurs.length;
                this.professeurs = [];
                this.affectations = [];
                this.afficherProfesseurs();
                this.updateStats();
                this.saveToLocalStorage();
                this.selectedProfessorIndex = null;
                
                const assignmentsTable = document.getElementById('assignmentsTableBody');
                if (assignmentsTable) assignmentsTable.innerHTML = '';
                const statsContainer = document.getElementById('assignmentsStats');
                if (statsContainer) statsContainer.innerHTML = '';
                const nonAffectesSection = document.getElementById('nonAffectesSection');
                if (nonAffectesSection) nonAffectesSection.remove();
                
                Swal.fire({
                    title: 'نجاح',
                    text: `تم حذف جميع الأساتذة (${count} أستاذ) بنجاح`,
                    icon: 'success',
                    timer: 3000,
                    showConfirmButton: true
                });
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
                    
                    const headers = jsonData[0];
                    let nomIndex = -1;
                    let matiereIndex = -1;
                    
                    headers.forEach((header, index) => {
                        const headerStr = String(header || '').trim();
                        if (headerStr.includes('اسم') || headerStr.includes('الاسم') || headerStr.includes('name') || headerStr.includes('Name')) {
                            nomIndex = index;
                        }
                        if (headerStr.includes('مادة') || headerStr.includes('المادة') || headerStr.includes('subject') || headerStr.includes('Subject') || headerStr.includes('matière')) {
                            matiereIndex = index;
                        }
                    });
                    
                    if (nomIndex === -1 && headers.length > 0) nomIndex = 0;
                    if (matiereIndex === -1 && headers.length > 1) matiereIndex = 1;
                    
                    if (nomIndex === -1 || matiereIndex === -1) {
                        Swal.fire('خطأ', 'الملف لا يحتوي على الأعمدة المطلوبة (الاسم الكامل، المادة)', 'error');
                        return;
                    }
                    
                    const nouveauxProfs = [];
                    for (let i = 1; i < jsonData.length; i++) {
                        const row = jsonData[i];
                        if (!row || row.length === 0) continue;
                        const nom = String(row[nomIndex] || '').trim();
                        const matiere = String(row[matiereIndex] || '').trim();
                        if (nom && matiere) {
                            nouveauxProfs.push({ nom, matiere, numero: '', indisponibilites: [] });
                        }
                    }
                    
                    if (this.professeurs.length > 0) {
                        Swal.fire({
                            title: 'خيارات الاستيراد',
                            html: `
                                <div style="text-align: right; margin: 20px 0;">
                                    <p>تم العثور على ${nouveauxProfs.length} أستاذ في الملف.</p>
                                    <p>اختر طريقة الاستيراد:</p>
                                    <div style="margin-top: 20px;">
                                        <button id="ajouterBtn" class="swal2-confirm swal2-styled" 
                                                style="background-color: #28a745; margin: 5px; width: 200px;">
                                            <i class="fas fa-plus"></i> إضافة إلى القائمة الحالية
                                        </button>
                                        <br>
                                        <button id="remplacerBtn" class="swal2-confirm swal2-styled" 
                                                style="background-color: #dc3545; margin: 5px; width: 200px;">
                                            <i class="fas fa-sync-alt"></i> استبدال القائمة الحالية
                                        </button>
                                    </div>
                                </div>
                            `,
                            showCancelButton: true,
                            cancelButtonText: 'إلغاء',
                            showConfirmButton: false,
                            allowOutsideClick: false,
                            didOpen: () => {
                                document.getElementById('ajouterBtn').addEventListener('click', () => {
                                    this.professeurs.push(...nouveauxProfs);
                                    this.finaliserImportExcel(nouveauxProfs.length, 'إضافة');
                                    Swal.close();
                                });
                                document.getElementById('remplacerBtn').addEventListener('click', () => {
                                    this.professeurs = nouveauxProfs;
                                    this.finaliserImportExcel(nouveauxProfs.length, 'استبدال');
                                    Swal.close();
                                });
                            }
                        });
                    } else {
                        this.professeurs = nouveauxProfs;
                        this.finaliserImportExcel(nouveauxProfs.length, 'إضافة');
                    }
                    
                } catch (error) {
                    Swal.fire('خطأ', `خطأ في قراءة الملف: ${error.message}`, 'error');
                }
            };
            reader.readAsArrayBuffer(file);
        };
        input.click();
    }

    finaliserImportExcel(count, operation) {
        this.afficherProfesseurs();
        this.updateStats();
        this.saveToLocalStorage();
        
        const matieres = this.obtenirListeMatieres();
        let message = `تم ${operation} ${count} أستاذ بنجاح!`;
        
        if (matieres.length > 0) {
            message += `\n\n📚 تم اكتشاف ${matieres.length} مادة:`;
            matieres.slice(0, 5).forEach(([matiere, count]) => {
                message += `\n• ${matiere} (${count} أستاذ)`;
            });
            if (matieres.length > 5) message += '\n...';
        }
        Swal.fire('نجاح', message, 'success');
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
            html: `هل أنت متأكد من حذف ${matieresArray.length} مادة من قائمة المواد فقط؟<br><br>
                   <strong>ملاحظة:</strong> لن يتم حذف هذه المواد من قائمة الأساتذة.<br><br>
                   ${matieresArray.map(m => `• ${m}`).join('<br>')}`,
            icon: 'warning',
            showCancelButton: true,
            confirmButtonText: 'نعم، احذف',
            cancelButtonText: 'إلغاء',
            confirmButtonColor: '#d33'
        }).then((result) => {
            if (result.isConfirmed) {
                const matieresASupprimer = new Set(matieresArray);
                this.matieres = this.matieres.filter(matiere => !matieresASupprimer.has(matiere.nom));
                document.getElementById('subjectCount').value = this.matieres.length;
                this.genererFormulaireMatieres();
                this.remplirFormulaireMatieres();
                this.saveToLocalStorage();
                this.closeModal('subjectsExtractModal');
                Swal.fire('نجاح', `تم حذف ${matieresArray.length} مادة من قائمة المواد فقط`, 'success');
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
                this.matieres = [];
                this.saveToLocalStorage();
                return;
            }
            
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
            
            if (this.matieres.length > nbrMatieres) {
                this.matieres = this.matieres.slice(0, nbrMatieres);
            } else {
                while (this.matieres.length < nbrMatieres) {
                    const i = this.matieres.length;
                    this.matieres.push({
                        nom: '',
                        date: `${(i % 30) + 1}/12/2024`,
                        periode: i % 2 === 0 ? 'صباح' : 'مساء',
                        heure_debut: i % 2 === 0 ? '08:00' : '14:00',
                        heure_fin: i % 2 === 0 ? '12:00' : '18:00',
                        duree: '04:00'
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
        // Stocker les infos des matières pour le rapport post-calcul
        this._matieresInfo = {};
        matieres.forEach(m => {
            this._matieresInfo[m.nom] = {
                date: m.date,
                periode: m.periode,
                demiJournee: `${m.date}-${m.periode}`
            };
        });

        matieres.forEach(matiere => {
            const demiJournee = `${matiere.date}-${matiere.periode}`;
            const dateHeure = `${matiere.date} ${matiere.heure_debut}-${matiere.heure_fin}`;

            this.salles.forEach(salle => {
                const profsAffectesASalle = [];
                const profsDisponibles = [...this.professeurs];

                for (let i = profsDisponibles.length - 1; i > 0; i--) {
                    const j = Math.floor(Math.random() * (i + 1));
                    [profsDisponibles[i], profsDisponibles[j]] = [profsDisponibles[j], profsDisponibles[i]];
                }

                for (let i = 0; i < nbrProfsSalle; i++) {
                    let profTrouve = null;

                    for (const prof of profsDisponibles) {
                        if (!profsAffectesASalle.includes(prof) &&
                            this.estProfesseurValide(prof, salle, demiJournee, matiere.nom, affectations, historiqueProfs)) {
                            profTrouve = prof;
                            break;
                        }
                    }

                    if (profTrouve) {
                        profsAffectesASalle.push(profTrouve);
                        const profId = profTrouve.nom;

                        if (!historiqueProfs[profId]) {
                            historiqueProfs[profId] = { salles: new Set(), matieres: new Set(), creneaux: new Set() };
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

        // ── Construire le rapport APRÈS que toutes les affectations sont finalisées ──
        // Cela garantit que historiqueProfs est complet et cohérent.
        this.rapportNonAffectes = {};

        const pasPropreMatiere = document.getElementById('constraint1').checked;
        const pasMemeSalle    = document.getElementById('constraint2').checked;
        const pasMemeGroupe   = document.getElementById('constraint3').checked;

        matieres.forEach(matiere => {
            const demiJournee = `${matiere.date}-${matiere.periode}`;
            const [date, periode] = [matiere.date, matiere.periode];

            // Ensemble des profs effectivement affectés à cette matière (dédupliqué)
            const profsAffectesSet = new Set(
                affectations.filter(a => a.matiere === matiere.nom).map(a => a.professeur)
            );

            this.rapportNonAffectes[matiere.nom] = {};

            this.professeurs.forEach(prof => {
                if (profsAffectesSet.has(prof.nom)) {
                    // Affecté — ne pas inclure dans le rapport des non-affectés
                    return;
                }

                // Ce prof n'a PAS été affecté à cette matière — calculer pourquoi
                const indispos = prof.indisponibilites || [];
                const profMatiere = typeof prof.matiere === 'string' ? prof.matiere.trim() : String(prof.matiere || '').trim();
                const hist = historiqueProfs[prof.nom];
                const raisons = [];

                if (indispos.some(([jour, p]) => jour === date && p === periode)) {
                    raisons.push('غير متوفر في هذا اليوم/الفترة');
                }
                if (pasPropreMatiere && profMatiere === matiere.nom) {
                    raisons.push('يدرّس هذه المادة (قيد 1)');
                }
                if (pasMemeGroupe && hist && hist.creneaux.has(demiJournee)) {
                    raisons.push('موزع في نفس الفترة على مادة أخرى (قيد 3)');
                }
                if (pasMemeSalle && hist && hist.salles.size >= this.salles.length) {
                    raisons.push('استنفد جميع القاعات (قيد 2)');
                }
                if (raisons.length === 0) {
                    raisons.push('لم تتوفر خانة شاغرة مناسبة');
                }

                this.rapportNonAffectes[matiere.nom][prof.nom] = {
                    prof,
                    raison: raisons.join(' | ')
                };
            });
        });

        return affectations;
    }

    estProfesseurValide(prof, salle, demiJournee, matiere, affectationsExistantes, historique) {
        const [date, periode] = demiJournee.split('-');
        
        const indispos = prof.indisponibilites || [];
        if (indispos.some(([jour, p]) => jour === date && p === periode)) return false;
        
        const pasPropreMatiere = document.getElementById('constraint1').checked;
        const pasMemeSalle = document.getElementById('constraint2').checked;
        const pasMemeGroupe = document.getElementById('constraint3').checked;
        
        const profId = prof.nom;
        const profMatiere = typeof prof.matiere === 'string' ? prof.matiere.trim() : String(prof.matiere || '').trim();
        
        if (pasPropreMatiere && profMatiere === matiere) return false;
        if (pasMemeSalle && historique[profId] && historique[profId].salles.has(salle)) return false;
        if (pasMemeGroupe && historique[profId] && historique[profId].creneaux.has(demiJournee)) return false;
        
        return true;
    }

    afficherResultats() {
        const tbody = document.getElementById('assignmentsTableBody');
        tbody.innerHTML = '';
        
        if (this.affectations.length === 0) return;
        
        const nbrProfsSalle = parseInt(document.getElementById('profsPerRoom').value) || 2;
        const affectationsParMatiere = {};
        
        this.affectations.forEach(affectation => {
            const matiere = affectation.matiere;
            if (!affectationsParMatiere[matiere]) affectationsParMatiere[matiere] = {};
            const salle = affectation.salle;
            if (!affectationsParMatiere[matiere][salle]) {
                affectationsParMatiere[matiere][salle] = { date_heure: affectation.date_heure, professeurs: [] };
            }
            affectationsParMatiere[matiere][salle].professeurs.push(affectation.professeur);
        });
        
        const thead = document.querySelector('#assignmentsTable thead tr');
        thead.innerHTML = '<th>الحالة</th>';
        for (let i = nbrProfsSalle; i >= 1; i--) {
            thead.innerHTML += `<th>الأستاذ ${i}</th>`;
        }
        thead.innerHTML += '<th>القاعة</th><th>تاريخ ووقت الامتحان</th><th>المادة</th>';
        
        for (const [matiere, salles] of Object.entries(affectationsParMatiere)) {
            for (const [salle, info] of Object.entries(salles)) {
                const tr = document.createElement('tr');
                let rowHTML = '';
                const statut = info.professeurs.length >= nbrProfsSalle ? '🟢 مكتمل' : '🟡 جزئي';
                rowHTML += `<td>${statut}</td>`;
                for (let i = nbrProfsSalle - 1; i >= 0; i--) {
                    rowHTML += `<td>${info.professeurs[i] || ''}</td>`;
                }
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
                <div class="stat-icon"><i class="${stat.icon}"></i></div>
                <h3>${stat.title}</h3>
                <p>${stat.value}</p>
            `;
            container.appendChild(card);
        });
    }

    afficherProfsNonAffectes() {
        if (this.affectations.length === 0) return;
        
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
            const assignmentsTable = document.querySelector('#affectations .table-container');
            assignmentsTable.parentNode.insertBefore(nonAffectesSection, assignmentsTable.nextSibling);
        }
        
        const container = document.getElementById('nonAffectesContainer');
        container.innerHTML = '';
        
        const profsAffectes = new Set(this.affectations.map(a => a.professeur));
        const matieresNonAffectes = {};
        const profsSansMatiere = [];
        
        this.professeurs.forEach(prof => {
            if (!profsAffectes.has(prof.nom)) {
                const matiere = prof.matiere ? prof.matiere.trim() : 'غير محدد';
                if (matiere && matiere !== 'غير محدد') {
                    if (!matieresNonAffectes[matiere]) matieresNonAffectes[matiere] = [];
                    matieresNonAffectes[matiere].push(prof);
                } else {
                    profsSansMatiere.push(prof);
                }
            }
        });
        
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
        
        Object.entries(matieresNonAffectes).forEach(([matiere, profs]) => {
            const matiereRow = document.createElement('tr');
            matiereRow.className = 'matiere-header-row';
            matiereRow.innerHTML = `<td colspan="5" style="background-color: #2a2a3c; font-weight: bold; color: #339af0;"><i class="fas fa-book"></i> ${matiere} (${profs.length} أستاذ)</td>`;
            tbody.appendChild(matiereRow);
            
            profs.forEach(prof => {
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
            
            const emptyRow = document.createElement('tr');
            emptyRow.innerHTML = '<td colspan="5" style="height: 10px;"></td>';
            tbody.appendChild(emptyRow);
        });
        
        if (profsSansMatiere.length > 0) {
            const sansMatiereRow = document.createElement('tr');
            sansMatiereRow.className = 'matiere-header-row';
            sansMatiereRow.innerHTML = `<td colspan="5" style="background-color: #2a2a3c; font-weight: bold; color: #fa5252;"><i class="fas fa-question-circle"></i> أساتذة بدون مادة محددة (${profsSansMatiere.length} أستاذ)</td>`;
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
        
        const totalNonAffectes = this.professeurs.length - profsAffectes.size;
        const summaryRow = document.createElement('tr');
        summaryRow.className = 'summary-row';
        summaryRow.innerHTML = `<td colspan="5" style="background-color: #495057; color: white; font-weight: bold; text-align: center;"><i class="fas fa-chart-bar"></i> إجمالي الأساتذة غير الموزعين: ${totalNonAffectes} من ${this.professeurs.length}</td>`;
        tbody.appendChild(summaryRow);
    }

    // ========== EXPORT EXCEL ==========

    showExcelConfigModal() {
        if (this.affectations.length === 0) {
            Swal.fire('تنبيه', 'يرجى توليد التوزيع أولاً', 'warning');
            return;
        }
        this.showModal('excelConfigModal');
    }

    /**
     * Returns professors NOT assigned to a specific subject (matiere).
     * A professor is considered "non-assigned for this subject" if:
     *  - They teach that subject (their prof.matiere matches), AND
     *  - They were never assigned to any room for that subject.
     * OR more broadly: any professor who does not appear in the assignments for that subject.
     * We use the broader definition: all profs whose name doesn't appear in affectations for this matiere.
     */
    /**
     * Retourne la liste des profs NON affectés à une matière donnée, avec raison.
     * Garantit : affectés (dédupliqués) + non-affectés = total profs.
     */
    obtenirProfsNonAffectesParMatiere(nomMatiere) {
        // Cas nominal : rapport disponible (généré dans la même session)
        if (this.rapportNonAffectes && this.rapportNonAffectes[nomMatiere]) {
            return Object.values(this.rapportNonAffectes[nomMatiere])
                .map(e => ({ ...e.prof, raison: e.raison }));
        }

        // Fallback (données rechargées depuis localStorage, rapport absent)
        const profsAffectesMatiere = new Set(
            this.affectations.filter(a => a.matiere === nomMatiere).map(a => a.professeur)
        );
        return this.professeurs
            .filter(prof => !profsAffectesMatiere.has(prof.nom))
            .map(prof => ({ ...prof, raison: 'غير محدد (بيانات مخزنة)' }));
    }

    genererExcel() {
        const university = document.getElementById('universityName').value.trim();
        const faculty    = document.getElementById('facultyName').value.trim();
        const ministere  = document.getElementById('faculty').value.trim();
        const academicYear = document.getElementById('academicYear').value.trim();
        const ecole      = document.getElementById('ecoleName').value.trim();

        if (!university || !faculty || !academicYear || !ecole) {
            Swal.fire('خطأ', 'يرجى ملء جميع الحقول الإلزامية', 'error');
            return;
        }

        try {
            const wb = XLSX.utils.book_new();
            const nbrProfsSalle = parseInt(document.getElementById('profsPerRoom').value) || 2;

            // ═══════════════════════════════════════════════════════════
            // PALETTE DE COULEURS PROFESSIONNELLE
            // ═══════════════════════════════════════════════════════════
            const CLR = {
                // En-têtes du document (bandeau institutionnel)
                INST_BG:   '1F3864',   // bleu marine foncé
                INST_FG:   'FFFFFF',   // blanc
                // Titre principal (nom de la feuille/matière)
                TITLE_BG:  '2E75B6',   // bleu moyen
                TITLE_FG:  'FFFFFF',
                // En-têtes de colonnes
                COL_BG:    '2E75B6',   // bleu moyen
                COL_FG:    'FFFFFF',
                // Lignes de données paires/impaires
                ROW_EVEN:  'DEEAF1',   // bleu très clair
                ROW_ODD:   'FFFFFF',   // blanc
                // Section non-affectés — titre
                NA_TITLE_BG: 'C55A11', // orange foncé
                NA_TITLE_FG: 'FFFFFF',
                // Section non-affectés — en-têtes colonnes
                NA_COL_BG:  'F4B942',  // orange clair
                NA_COL_FG:  '000000',
                // Section non-affectés — lignes données
                NA_ROW_EVEN: 'FFF2CC', // jaune très clair
                NA_ROW_ODD:  'FFFFFF',
                // Ligne de résumé / total
                TOTAL_BG:   '1F3864',
                TOTAL_FG:   'FFFFFF',
                // Message "tout affecté"
                OK_BG:      '70AD47',
                OK_FG:      'FFFFFF',
                // Bordures
                BORDER_DARK: '1F3864',
                BORDER_MED:  '2E75B6',
                BORDER_LIGHT:'BDD7EE',
            };

            // ═══════════════════════════════════════════════════════════
            // HELPER : fabrique un objet style complet
            // ═══════════════════════════════════════════════════════════
            const S = ({
                bg = null, fg = '000000', sz = 11, bold = false,
                italic = false, halign = 'center', valign = 'center',
                wrapText = true, borderType = 'all', borderColor = null,
                topBorderStyle = 'thin', bottomBorderStyle = 'thin'
            } = {}) => {
                const bc = borderColor || CLR.BORDER_LIGHT;
                const border = {};
                if (borderType === 'all' || borderType === 'outer') {
                    const mk = (style, color) => ({ style, color: { rgb: color } });
                    border.top    = mk(topBorderStyle,    bc);
                    border.bottom = mk(bottomBorderStyle, bc);
                    border.left   = mk('thin', bc);
                    border.right  = mk('thin', bc);
                }
                const s = {
                    font: { name: 'Arial', sz, bold, italic, color: { rgb: fg } },
                    alignment: { horizontal: halign, vertical: valign,
                                 readingOrder: 2, wrapText },
                    border
                };
                if (bg) s.fill = { patternType: 'solid', fgColor: { rgb: bg } };
                return s;
            };

            // ═══════════════════════════════════════════════════════════
            // HELPER : applique un style à toute une plage
            // ═══════════════════════════════════════════════════════════
            const applyRange = (ws, r1, c1, r2, c2, style) => {
                for (let R = r1; R <= r2; R++) {
                    for (let C = c1; C <= c2; C++) {
                        const ref = XLSX.utils.encode_cell({ r: R, c: C });
                        if (!ws[ref]) ws[ref] = { v: '', t: 's' };
                        ws[ref].s = style;
                    }
                }
            };

            // ═══════════════════════════════════════════════════════════
            // HELPER : fusionne des cellules et place la valeur + style
            // ═══════════════════════════════════════════════════════════
            const merge = (ws, r1, c1, r2, c2, value, style) => {
                if (!ws['!merges']) ws['!merges'] = [];
                ws['!merges'].push({ s: { r: r1, c: c1 }, e: { r: r2, c: c2 } });
                const ref = XLSX.utils.encode_cell({ r: r1, c: c1 });
                ws[ref] = { v: value, t: 's', s: style };
                // Appliquer le même style aux cellules fusionnées (pour les bordures)
                for (let R = r1; R <= r2; R++) {
                    for (let C = c1; C <= c2; C++) {
                        const cellRef = XLSX.utils.encode_cell({ r: R, c: C });
                        if (!ws[cellRef]) ws[cellRef] = { v: '', t: 's' };
                        ws[cellRef].s = style;
                    }
                }
            };

            // ═══════════════════════════════════════════════════════════
            // HELPER : hauteurs de lignes
            // ═══════════════════════════════════════════════════════════
            const setRowHeights = (ws, heights) => {
                if (!ws['!rows']) ws['!rows'] = [];
                heights.forEach(([r, h]) => {
                    while (ws['!rows'].length <= r) ws['!rows'].push({});
                    ws['!rows'][r] = { hpt: h };
                });
            };

            // ═══════════════════════════════════════════════════════════
            // STYLES PRÉDÉFINIS
            // ═══════════════════════════════════════════════════════════
            const ST = {
                inst:     S({ bg: CLR.INST_BG,     fg: CLR.INST_FG,     sz: 11, bold: false, borderColor: CLR.BORDER_DARK }),
                instBold: S({ bg: CLR.INST_BG,     fg: CLR.INST_FG,     sz: 13, bold: true,  borderColor: CLR.BORDER_DARK }),
                title:    S({ bg: CLR.TITLE_BG,    fg: CLR.TITLE_FG,    sz: 14, bold: true,  borderColor: CLR.BORDER_DARK, topBorderStyle: 'medium', bottomBorderStyle: 'medium' }),
                colHead:  S({ bg: CLR.COL_BG,      fg: CLR.COL_FG,      sz: 11, bold: true,  borderColor: CLR.BORDER_MED }),
                rowEven:  S({ bg: CLR.ROW_EVEN,    fg: '000000',        sz: 11, borderColor: CLR.BORDER_LIGHT }),
                rowOdd:   S({ bg: CLR.ROW_ODD,     fg: '000000',        sz: 11, borderColor: CLR.BORDER_LIGHT }),
                naTitle:  S({ bg: CLR.NA_TITLE_BG, fg: CLR.NA_TITLE_FG, sz: 12, bold: true,  borderColor: CLR.BORDER_DARK, topBorderStyle: 'medium', bottomBorderStyle: 'medium' }),
                naColHead:S({ bg: CLR.NA_COL_BG,   fg: CLR.NA_COL_FG,   sz: 11, bold: true,  borderColor: '000000' }),
                naEven:   S({ bg: CLR.NA_ROW_EVEN, fg: '000000',        sz: 10, borderColor: '999999' }),
                naOdd:    S({ bg: CLR.NA_ROW_ODD,  fg: '000000',        sz: 10, borderColor: '999999' }),
                total:    S({ bg: CLR.TOTAL_BG,    fg: CLR.TOTAL_FG,    sz: 11, bold: true,  borderColor: CLR.BORDER_DARK, topBorderStyle: 'medium', bottomBorderStyle: 'medium' }),
                ok:       S({ bg: CLR.OK_BG,       fg: CLR.OK_FG,       sz: 11, bold: true,  borderColor: '70AD47' }),
            };

            // ═══════════════════════════════════════════════════════════
            // HELPER PRINCIPAL : construit un worksheet complet
            // params:
            //   headerLines  : [{label, value}]  → lignes institutionnelles
            //   tableTitle   : string            → titre principal (matière, etc.)
            //   columns      : string[]          → noms des colonnes
            //   rows         : any[][]           → données
            //   colWidths    : number[]          → largeurs colonnes (wch)
            //   naSection    : { title, columns, rows } | null
            // ═══════════════════════════════════════════════════════════
            const buildSheet = (headerLines, tableTitle, columns, rows, colWidths, naSection = null) => {
                const ws = {};
                const nbCols = Math.max(columns.length, naSection ? (naSection.columns || []).length : 0);
                let r = 0;

                // ── Lignes institutionnelles ──────────────────────────
                const rowHeights = [];
                headerLines.forEach((line, i) => {
                    const isFirst = (i === 0);
                    const st = isFirst ? ST.instBold : ST.inst;
                    const text = line.value ? `${line.label} ${line.value}` : line.label;
                    merge(ws, r, 0, r, nbCols - 1, text, st);
                    rowHeights.push([r, isFirst ? 22 : 18]);
                    r++;
                });

                // ── Ligne vide de séparation ──────────────────────────
                rowHeights.push([r, 8]);
                r++;

                // ── Titre principal ───────────────────────────────────
                merge(ws, r, 0, r, nbCols - 1, tableTitle, ST.title);
                rowHeights.push([r, 28]);
                r++;

                // ── Ligne vide ────────────────────────────────────────
                rowHeights.push([r, 8]);
                r++;

                // ── En-têtes colonnes ─────────────────────────────────
                const dataStartRow = r + 1; // pour les merges de feuilles globales
                columns.forEach((col, c) => {
                    const ref = XLSX.utils.encode_cell({ r, c });
                    ws[ref] = { v: col, t: 's', s: ST.colHead };
                });
                rowHeights.push([r, 22]);
                const tableHeaderRow = r;
                r++;

                // ── Lignes de données ─────────────────────────────────
                rows.forEach((row, idx) => {
                    const st = idx % 2 === 0 ? ST.rowEven : ST.rowOdd;
                    row.forEach((cell, c) => {
                        const ref = XLSX.utils.encode_cell({ r, c });
                        ws[ref] = { v: cell == null ? '' : String(cell), t: 's', s: st };
                    });
                    // Remplir les cellules manquantes
                    for (let c = row.length; c < columns.length; c++) {
                        const ref = XLSX.utils.encode_cell({ r, c });
                        ws[ref] = { v: '', t: 's', s: st };
                    }
                    rowHeights.push([r, 18]);
                    r++;
                });

                // ── Section non-affectés ──────────────────────────────
                if (naSection) {
                    // Ligne vide de séparation
                    rowHeights.push([r, 10]);
                    r++;

                    // Titre section
                    merge(ws, r, 0, r, nbCols - 1, naSection.title, ST.naTitle);
                    rowHeights.push([r, 24]);
                    r++;

                    if (naSection.rows.length === 0) {
                        // Message "tout affecté"
                        merge(ws, r, 0, r, nbCols - 1, '✅ جميع الأساتذة تم توزيعهم على هذه المادة', ST.ok);
                        rowHeights.push([r, 20]);
                        r++;
                    } else {
                        // En-têtes colonnes non-affectés
                        const naCols = naSection.columns;
                        naCols.forEach((col, c) => {
                            const ref = XLSX.utils.encode_cell({ r, c });
                            ws[ref] = { v: col, t: 's', s: ST.naColHead };
                        });
                        // Remplir les cellules manquantes jusqu'à nbCols
                        for (let c = naCols.length; c < nbCols; c++) {
                            const ref = XLSX.utils.encode_cell({ r, c });
                            ws[ref] = { v: '', t: 's', s: ST.naColHead };
                        }
                        rowHeights.push([r, 20]);
                        r++;

                        // Données non-affectés
                        naSection.rows.forEach((row, idx) => {
                            const st = idx % 2 === 0 ? ST.naEven : ST.naOdd;
                            row.forEach((cell, c) => {
                                const ref = XLSX.utils.encode_cell({ r, c });
                                ws[ref] = { v: cell == null ? '' : String(cell), t: 's', s: st };
                            });
                            for (let c = row.length; c < nbCols; c++) {
                                const ref = XLSX.utils.encode_cell({ r, c });
                                ws[ref] = { v: '', t: 's', s: st };
                            }
                            rowHeights.push([r, 18]);
                            r++;
                        });

                        // Ligne de total
                        merge(ws, r, 0, r, nbCols - 1, naSection.totalLine, ST.total);
                        rowHeights.push([r, 22]);
                        r++;
                    }
                }

                // ── Dimensions ───────────────────────────────────────
                ws['!ref'] = XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: r - 1, c: nbCols - 1 } });
                ws['!cols'] = colWidths.map(w => ({ wch: w }));
                ws['!rtl'] = true;
                setRowHeights(ws, rowHeights);

                return ws;
            };

            // ═══════════════════════════════════════════════════════════
            // Lignes institutionnelles communes
            // ═══════════════════════════════════════════════════════════
            const instLines = [
                { label: ministere },
                { label: 'الأكادمية:', value: faculty },
                { label: 'المديرية:', value: university },
                { label: 'المؤسسة:', value: ecole },
                { label: 'السنة الدراسية:', value: academicYear },
                { label: 'نوع الامتحان:', value: this.type_examen },
            ];

            // ═══════════════════════════════════════════════════════════
            // FEUILLE 1 : الجدول العام
            // ═══════════════════════════════════════════════════════════
            const globalCols = ['المادة', 'تاريخ ووقت الامتحان', 'القاعة'];
            for (let i = nbrProfsSalle; i >= 1; i--) globalCols.push(`الأستاذ ${i}`);

            const grouped = {};
            this.affectations.forEach(a => {
                const key = `${a.matiere}_${a.salle}`;
                if (!grouped[key]) grouped[key] = { matiere: a.matiere, salle: a.salle, date: a.date_heure, profs: [] };
                grouped[key].profs.push(a.professeur);
            });

            const globalRows = Object.values(grouped).map(item => {
                const row = [item.matiere, item.date, item.salle];
                for (let i = nbrProfsSalle - 1; i >= 0; i--) row.push(item.profs[i] || '');
                return row;
            });

            const globalColW = [28, 28, 14];
            for (let i = 0; i < nbrProfsSalle; i++) globalColW.push(28);

            const wsGlobal = buildSheet(instLines, 'الجدول العام لتوزيع الأساتذة', globalCols, globalRows, globalColW);
            XLSX.utils.book_append_sheet(wb, wsGlobal, 'الجدول العام');

            // ═══════════════════════════════════════════════════════════
            // FEUILLE 2 : غير الموزعين (global)
            // ═══════════════════════════════════════════════════════════
            const profsAffectesGlobal = new Set(this.affectations.map(a => a.professeur));
            const naGlobalCols = ['الاسم الكامل', 'المادة', 'رقم التأجير', 'سبب عدم التوزيع'];
            const naGlobalRows = [];
            const naGlobalColW = [30, 20, 14, 35];

            // Grouper par matière pour un affichage structuré
            const naParMatiere = {};
            this.professeurs.forEach(prof => {
                if (!profsAffectesGlobal.has(prof.nom)) {
                    const mat = prof.matiere || 'غير محدد';
                    if (!naParMatiere[mat]) naParMatiere[mat] = [];
                    naParMatiere[mat].push(prof);
                }
            });

            Object.entries(naParMatiere).forEach(([mat, profs]) => {
                // Ligne de séparateur matière (simulé en tant que données en couleur)
                naGlobalRows.push([`── ${mat} (${profs.length} أستاذ) ──`, '', '', '']);
                profs.forEach(prof => {
                    naGlobalRows.push([
                        prof.nom,
                        prof.matiere || '',
                        prof.numero || '',
                        prof.indisponibilites && prof.indisponibilites.length > 0 ? 'لديه قيود توفر' : 'لم تتوفر خانة شاغرة'
                    ]);
                });
            });

            const naTotalLine = `إجمالي الأساتذة: ${this.professeurs.length}  |  الموزعون: ${profsAffectesGlobal.size}  |  غير الموزعين: ${this.professeurs.length - profsAffectesGlobal.size}`;

            const wsNA = buildSheet(
                instLines,
                'قائمة الأساتذة غير الموزعين',
                naGlobalCols,
                naGlobalRows,
                naGlobalColW,
                null  // pas de double section ici, on affiche tout dans le tableau principal
            );

            // Ajouter la ligne de total manuellement en fin
            const naRange = XLSX.utils.decode_range(wsNA['!ref']);
            const totalR = naRange.e.r + 1;
            merge(wsNA, totalR, 0, totalR, naGlobalCols.length - 1, naTotalLine, ST.total);
            wsNA['!ref'] = XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: totalR, c: naGlobalCols.length - 1 } });
            if (!wsNA['!rows']) wsNA['!rows'] = [];
            while (wsNA['!rows'].length <= totalR) wsNA['!rows'].push({});
            wsNA['!rows'][totalR] = { hpt: 22 };

            XLSX.utils.book_append_sheet(wb, wsNA, 'غير الموزعين');

            // ═══════════════════════════════════════════════════════════
            // FEUILLES PAR MATIÈRE
            // ═══════════════════════════════════════════════════════════
            const matieresGroup = {};
            this.affectations.forEach(a => {
                if (!matieresGroup[a.matiere]) matieresGroup[a.matiere] = [];
                matieresGroup[a.matiere].push(a);
            });

            for (const [matiere, affectationsMatiere] of Object.entries(matieresGroup)) {
                // Colonnes du tableau d'affectation
                const matCols = ['القاعة', 'تاريخ ووقت الامتحان'];
                for (let i = nbrProfsSalle; i >= 1; i--) matCols.push(`الأستاذ ${i}`);

                // Données du tableau d'affectation
                const groupedBySalle = {};
                affectationsMatiere.forEach(a => {
                    if (!groupedBySalle[a.salle]) groupedBySalle[a.salle] = { salle: a.salle, date: a.date_heure, profs: [] };
                    groupedBySalle[a.salle].profs.push(a.professeur);
                });

                const matRows = Object.values(groupedBySalle).map(item => {
                    const row = [item.salle, item.date];
                    for (let i = nbrProfsSalle - 1; i >= 0; i--) row.push(item.profs[i] || '');
                    return row;
                });

                // Largeurs colonnes
                const matColW = [14, 26];
                for (let i = 0; i < nbrProfsSalle; i++) matColW.push(28);
                // Ajouter des colonnes pour la section non-affectés (min 5 colonnes)
                while (matColW.length < 5) matColW.push(28);
                matColW.push(14, 30); // رقم التأجير + سبب

                // Section non-affectés
                const profsNonAff = this.obtenirProfsNonAffectesParMatiere(matiere);
                const nbAff = new Set(this.affectations.filter(a => a.matiere === matiere).map(a => a.professeur)).size;
                const nbNon = profsNonAff.length;
                const total = this.professeurs.length;

                const naRows = profsNonAff.map(prof => {
                    const indispoText = prof.indisponibilites && prof.indisponibilites.length > 0
                        ? prof.indisponibilites.map(([j, p]) => `${j}-${p === 'matin' ? 'ص' : 'م'}`).join(' | ')
                        : 'لا يوجد';
                    return [prof.nom, prof.matiere || '', prof.numero || '', indispoText, prof.raison || 'غير محدد'];
                });

                const naCols   = ['الاسم الكامل', 'مادته', 'رقم التأجير', 'عدم التوفر', 'سبب عدم التوزيع'];
                const naColW   = [28, 20, 12, 22, 32];
                const nbCols   = Math.max(matCols.length, naCols.length);
                // Fusionner les largeurs
                const finalColW = [];
                for (let i = 0; i < nbCols; i++) {
                    finalColW.push(Math.max(matColW[i] || 14, naColW[i] || 14));
                }

                const wsMatiere = buildSheet(
                    instLines,
                    `جدول مراقبة مادة : ${matiere}`,
                    matCols,
                    matRows,
                    finalColW,
                    {
                        title:     `الأساتذة غير الموزعين على هذه المادة`,
                        columns:   naCols,
                        rows:      naRows,
                        totalLine: `✔ الموزعون: ${nbAff}  ✖ غير الموزعين: ${nbNon}  Σ الإجمالي: ${total}  (${nbAff} + ${nbNon} = ${total})`
                    }
                );

                const sheetName = matiere.substring(0, 31);
                XLSX.utils.book_append_sheet(wb, wsMatiere, sheetName);
            }

            // ═══════════════════════════════════════════════════════════
            // SAUVEGARDE
            // ═══════════════════════════════════════════════════════════
            const fileName = `توزيع_الأساتذة_${new Date().toISOString().slice(0,10)}.xlsx`;
            XLSX.writeFile(wb, fileName);

            this.closeModal('excelConfigModal');
            Swal.fire({
                title: 'نجاح',
                html: `تم إنشاء ملف Excel بنجاح يحتوي على:<br>
                       <b>1.</b> الجدول العام<br>
                       <b>2.</b> قائمة الأساتذة غير الموزعين<br>
                       <b>3.</b> جداول لكل مادة مع قائمة غير الموزعين وأسباب الإقصاء`,
                icon: 'success'
            });

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

document.addEventListener('DOMContentLoaded', () => {
    window.app = new ApplicationAffectation();
});