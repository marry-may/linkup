import { Button } from "@/components/ui/button";

export default function Home() {
  return (
    <main className="flex flex-1 flex-col items-center justify-center gap-6 px-4 py-16 text-center">
      <h1 className="text-4xl font-semibold tracking-tight">LinkUp</h1>
      <p className="max-w-sm text-muted-foreground">
        See who&apos;s at the event, find the people you should meet, and connect.
      </p>
      <Button size="lg">Coming soon</Button>
    </main>
  );
}
