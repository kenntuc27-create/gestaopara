import { Link, useLocation, useNavigate, useRouter } from "@tanstack/react-router";
import {
  LayoutDashboard,
  Users,
  Settings,
  LogOut,
  ShieldCheck,
  Loader2,
  ArrowLeft,
  PanelLeftClose,
  PanelLeftOpen,
  Menu,
  Calculator,
  ScrollText,
  TrendingUp,
  Briefcase,
  Target,
  Trophy,
  Clock,
  CalendarCheck,
  UserCog,
  FileSignature,
  Bell,
  GraduationCap,
  History,
  Activity,
  Gavel,
  Building2,
  FileCheck,
} from "lucide-react";
import { ReactNode, useEffect, useState } from "react";
import { cn } from "@/lib/utils";
import logo from "@/assets/logo.png";
import { InstallAppButton } from "@/components/InstallAppButton";
import { useAuth } from "@/hooks/useAuth";
import { Button } from "@/components/ui/button";
import { DailyScorePopup } from "@/components/DailyScorePopup";
import { Sheet, SheetContent, SheetTrigger } from "@/components/ui/sheet";
import { NotificationsBootstrap } from "@/components/NotificationsBootstrap";
import { SidebarNavGroup, type NavItem } from "@/components/SidebarGroup";

const OPERATIONAL_PREFIXES = ["/", "/novo", "/historico", "/fornecedores", "/configuracoes", "/edital", "/central"];

function requiresOperational(pathname: string) {
  if (pathname === "/") return true;
  return OPERATIONAL_PREFIXES.some((p) => p !== "/" && pathname.startsWith(p));
}

interface NavGroup {
  id: string;
  label: string;
  items: NavItem[];
}

export function AppShell({ children, title, actions }: { children: ReactNode; title?: string; actions?: ReactNode }) {
  const loc = useLocation();
  const navigate = useNavigate();
  const router = useRouter();
  const { user, loading, isAdmin, signOut, roles, hasOperationalAccess, sectorName } = useAuth();

  const [sidebarOpen, setSidebarOpen] = useState(true);
  const [sidebarCollapsed, setSidebarCollapsed] = useState(false);
  const [mobileOpen, setMobileOpen] = useState(false);
  const [mustChange, setMustChange] = useState(false);

  // THEME (corrigido - sem duplicação)
  const [darkMode, setDarkMode] = useState(false);

  useEffect(() => {
    const savedTheme = localStorage.getItem("theme");

    if (savedTheme === "dark") {
      document.documentElement.classList.add("dark");
      setDarkMode(true);
    }
  }, []);

  const toggleTheme = () => {
    const html = document.documentElement;

    if (html.classList.contains("dark")) {
      html.classList.remove("dark");
      localStorage.setItem("theme", "light");
      setDarkMode(false);
    } else {
      html.classList.add("dark");
      localStorage.setItem("theme", "dark");
      setDarkMode(true);
    }
  };

  useEffect(() => {
    if (!loading && !user) navigate({ to: "/login" });
  }, [loading, user, navigate]);

  if (loading || !user) {
    return (
      <div className="min-h-screen flex items-center justify-center">
        <Loader2 className="size-6 animate-spin text-muted-foreground" />
      </div>
    );
  }

  const sidebarWidth = sidebarCollapsed ? "w-14" : "w-60";

  const groups: NavGroup[] = [];

  if (hasOperationalAccess) {
    groups.push({
      id: "dashboard",
      label: "Dashboard",
      items: [{ to: "/", label: "Visão Geral", icon: LayoutDashboard, exact: true }],
    });
  }

  groups.push({
    id: "performance",
    label: "Performance",
    items: [
      { to: "/equipe/performance-geral", label: "Dashboard Geral", icon: TrendingUp },
      { to: "/equipe/performance", label: "Performance", icon: Activity },
      { to: "/equipe/metas", label: "Metas", icon: Target },
    ],
  });

  const NavList = ({ onNavigate, collapsed = false }: any) => (
    <nav className="flex-1 py-2 space-y-1 overflow-y-auto">
      {groups.map((g) => (
        <SidebarNavGroup
          key={g.id}
          id={g.id}
          label={g.label}
          items={g.items}
          collapsedSidebar={collapsed}
          onNavigate={onNavigate}
        />
      ))}
    </nav>
  );

  const SidebarFooter = ({ collapsed = false }: any) => (
    <div className={cn("p-3 space-y-2 border-t border-sidebar-border", collapsed && "px-1.5")}>
      <Button
        size={collapsed ? "icon" : "sm"}
        variant="ghost"
        onClick={async () => {
          await signOut();
          navigate({ to: "/login" });
        }}
      >
        <LogOut className="size-4" />
        {!collapsed && <span className="ml-2">Sair</span>}
      </Button>
      {!collapsed && <InstallAppButton />}
    </div>
  );

  return (
    <div className="flex min-h-screen bg-background">
      {/* SIDEBAR */}
      {sidebarOpen && (
        <aside className={cn("hidden md:flex flex-col border-r", sidebarWidth)}>
          <div className="p-4 border-b">
            <img src={logo} className="h-10" />
          </div>

          <NavList collapsed={sidebarCollapsed} />
          <SidebarFooter collapsed={sidebarCollapsed} />
        </aside>
      )}

      {/* MAIN */}
      <div className="flex-1 flex flex-col">
        <header className="h-14 flex items-center justify-between border-b px-4">
          <div className="flex items-center gap-2">
            <Button size="icon" onClick={() => setSidebarCollapsed(v => !v)}>
              <Menu />
            </Button>

            <h1 className="font-semibold">{title}</h1>
          </div>

          <div>{actions}</div>
        </header>

        <main className="flex-1 p-4">{children}</main>
      </div>

      <DailyScorePopup />
      <NotificationsBootstrap />
    </div>
  );
}