import { createFileRoute } from "@tanstack/react-router";
import { Button } from "@/components/ui/button";

export const Route = createFileRoute("/")({
  head: () => ({
    meta: [
      { title: "SchoolBot Mail — votre boîte mail triée par IA, sur votre PC" },
      { name: "description", content: "Logiciel gratuit pour directions et secrétariats d'écoles : tri des mails et réponses proposées par IA, sans que vos mails quittent l'ordinateur." },
      { property: "og:title", content: "SchoolBot Mail — tri des mails par IA, en local" },
      { property: "og:description", content: "Gratuit, installé sur votre PC, vos mails ne le quittent jamais. Offert par SchoolBot.be." },
      { property: "og:type", content: "website" },
      { name: "twitter:card", content: "summary" },
    ],
  }),
  component: Index,
});

const steps = [
  ["Installez Python", "Gratuit, depuis python.org (cochez « Add to PATH »)."],
  ["Installez l'IA locale", "Ollama (ollama.com), puis la commande : ollama pull mistral-small. Le programme choisit le meilleur modèle installé."],
  ["Lancez SchoolBot Mail", "Double-cliquez sur le fichier : le navigateur s'ouvre tout seul."],
  ["Tapez votre adresse", "Les serveurs sont trouvés automatiquement. Ajoutez le mot de passe, testez, c'est parti."],
];

async function download() {
  // Le fichier est récupéré par la page puis enregistré localement : évite le blocage
  // des téléchargements directs dans l'aperçu protégé.
  const res = await fetch("/schoolbot-mail.py", { credentials: "include" });
  const blob = new Blob([await res.text()], { type: "application/octet-stream" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = "schoolbot-mail.py";
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

function Index() {
  return (
    <main className="min-h-screen bg-background text-foreground">
      <section className="mx-auto max-w-3xl px-6 py-20">
        <p className="text-sm text-muted-foreground">Offert par SchoolBot.be</p>
        <h1 className="mt-2 text-4xl font-bold tracking-tight">SchoolBot Mail</h1>
        <p className="mt-4 text-lg text-muted-foreground">
          Chaque matin, vos mails du jour triés en Urgent, À traiter, À lire, Postposable — avec une réponse
          déjà rédigée. Tout se passe sur votre ordinateur : aucun mail n'en sort.
        </p>
        <div className="mt-8 flex flex-wrap gap-3">
          <Button size="lg" onClick={download}>Télécharger (gratuit)</Button>
          <Button asChild size="lg" variant="outline">
            <a href="https://schoolbot.be" target="_blank" rel="noreferrer">Découvrir SchoolBot</a>
          </Button>
        </div>
        <ol className="mt-14 space-y-4">
          {steps.map(([t, d], i) => (
            <li key={t} className="flex gap-4 rounded-lg border bg-card p-4">
              <span className="font-bold text-primary">{i + 1}</span>
              <div>
                <p className="font-semibold">{t}</p>
                <p className="text-sm text-muted-foreground">{d}</p>
              </div>
            </li>
          ))}
        </ol>
      </section>
    </main>
  );
}
