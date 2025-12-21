/**
 * nav.js
 * Berisi fungsi untuk mengontrol navigasi sidebar.
 */

function openNav() {
    document.getElementById("mySidebar").style.width = "250px";
    document.getElementById("main").style.marginLeft = "250px";
    document.querySelector(".hamburger-menu").style.display = 'none';
}

function closeNav() {
    document.getElementById("mySidebar").style.width = "0";
    document.getElementById("main").style.marginLeft = "0";
    document.querySelector(".hamburger-menu").style.display = 'block';
}
