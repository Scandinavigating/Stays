@echo off
echo Genereren van nieuwe index.html...
python Conversion_Page.py

echo Uploaden naar GitHub Pages...
git add .
git commit -m "Auto-update conversion grid"
git push origin main

echo Klaar! Je live link is over 30 seconden bijgewerkt.
pause
```[cite: 2]

---

### Jouw toekomstige routine bij nieuwe video's:

1. Sleep de gemaakte omslagfoto's in de map `Thumbnails`[cite: 2].
2. **Dubbelklik op `update_site.bat`:**
   * Python leest Excel uit[cite: 2].
   * De nieuwe tegels worden toegevoegd aan `index.html`[cite: 2].
   * De wijzigingen worden direct doorgestuurd naar GitHub[cite: 7].
   * GitHub Pages bouwt de site automatisch opnieuw op de achtergrond[cite: 7, 8].

Je hoeft de link in je Instagram- of Facebook-bio nooit meer aan te passen; kijkers zien bij het openen van je bio direct de nieuwste tegels bovenaan staan[cite: 1, 2, 7].