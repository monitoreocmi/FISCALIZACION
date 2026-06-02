
        function volverAtras() { window.history.back(); }

        function openModal(src) {
            const modal = document.getElementById('myModal');
            const img = document.getElementById('imgModal');
            if(modal && img) { modal.style.display = "flex"; img.src = src; }
        }

        function cambiarMes() {
            let selectedValue = document.getElementById("mes-selector").value;
            document.querySelectorAll(".mes-container").forEach(e => e.style.display = "none");
            let targetMes = document.getElementById(selectedValue);
            if(targetMes) { targetMes.style.display = "block"; switchTab('inc'); }
        }

        function switchTab(tipo) {
            let selector = document.getElementById("mes-selector");
            if(!selector) return;
            let mesNombre = selector.value.split("-")[1];
            document.querySelectorAll(".tab-btn").forEach(b => b.classList.remove("active"));
            if(tipo === 'inc') {
                document.querySelectorAll(".tab-btn")[0].classList.add("active");
                document.getElementById("content-inc-" + mesNombre).style.display = "block";
                document.getElementById("content-cob-" + mesNombre).style.display = "none";
            } else {
                document.querySelectorAll(".tab-btn")[1].classList.add("active");
                document.getElementById("content-inc-" + mesNombre).style.display = "none";
                document.getElementById("content-cob-" + mesNombre).style.display = "block";
            }
        }
        