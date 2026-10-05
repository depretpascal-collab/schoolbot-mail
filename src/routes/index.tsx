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
  ["Téléchargez et décompressez", "Récupérez le fichier zip, puis décompressez-le (clic droit → « Extraire tout »)."],
  ["Double-cliquez sur SchoolBot Mail", "Windows peut afficher un écran bleu de sécurité : cliquez sur « Informations complémentaires », puis « Exécuter quand même »."],
  ["Laissez l'assistant travailler", "L'IA s'installe et se prépare toute seule, sans aucune commande à taper."],
  ["Tapez votre adresse", "Les serveurs sont trouvés automatiquement. Ajoutez le mot de passe, testez, c'est parti."],
];

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
          <Button asChild size="lg">
            <a href="https://github.com/depretpascal-collab/schoolbot-mail/releases/latest/download/SchoolBot-Mail-Windows.zip">Télécharger pour Windows (gratuit)</a>
          </Button>
          <Button asChild size="lg" variant="outline">
            <a href="https://github.com/depretpascal-collab/schoolbot-mail/releases/latest/download/SchoolBot-Mail-Mac.zip">Pour Mac</a>
          </Button>
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
