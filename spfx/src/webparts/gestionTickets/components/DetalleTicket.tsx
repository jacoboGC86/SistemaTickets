import * as React from 'react';
import {
  Panel,
  PanelType,
  Stack,
  TextField,
  Dropdown,
  IDropdownOption,
  PrimaryButton,
  DefaultButton,
  IStackTokens,
  MessageBar,
  MessageBarType,
  Spinner,
  SpinnerSize,
  Label,
  Separator,
  IconButton,
  TooltipHost,
  Callout,
  DirectionalHint,
} from '@fluentui/react';
import ListSvc from '../../../services/ListSvc';
import UserSvc from '../../../services/UserSvc';
import EvaluacionRiesgo, { IEvaluacionRiesgoValues } from './EvaluacionRiesgo';

// ─── Types ─────────────────────────────────────────────────────────────────────

interface IApprovalStep {
  stepName: string;
  useUserDepartment?: boolean;
  approvalGroup?: string | null;
}

interface IFormField {
  tag: string;
  id?: string;
  text?: string;
  checked?: boolean;
  value?: string;
}

interface ITemplate {
  approvalPath: IApprovalStep[];
  formTemplate?: IFormField[];
  processManager?: { id: number };
}

interface IAprobacionRow {
  id: number;
  title: string;
  resultado: string;
  responsable: string | null;
  responsableEmail: string | null;
  modified: string | null;
  grupoId: number | null;
}

interface IResolvedStep {
  stepName: string;
  responsable: string | null;
  responsableEmail: string | null;
}

interface IAdjunto {
  name: string;
  url: string;
}

interface IComment {
  id: number;
  author: string;
  text: string;
  date: string;
  adjuntos: { name: string; url: string }[];
}

interface ITicketData {
  id: number;
  title: string;
  solicitante: string;
  solicitanteId: number;
  fechaRegistro: string;
  tipoTicket: string;
  categoriaId: number | null;
  categoriaTitle: string;
  descripcion: string;
  departamento: string;
  planta: string;
  aNombreDe: string | null;
  prioridad: string;
  status: string;
  templateConfiguracion: ITemplate | null;
  solucionDetallada: string;
  reaperturado: boolean;
  motivoReapertura: string;
  atencion: boolean;
  impactoSeguridadInformacion: string;
  descripcionImpactoSeguridad: string;
  processManagerId: number;
  processManagerName: string;
  managerId: number | null;
  managerName: string | null;
  managerEmail: string | null;
  modified: string;
}

// ─── Props ─────────────────────────────────────────────────────────────────────

export interface IDetalleTicketProps {
  isOpen: boolean;
  onDismiss: () => void;
  ticketId?: number | null;
}

// ─── Constants ─────────────────────────────────────────────────────────────────

const stackTokens: IStackTokens = { childrenGap: 12 };

const typeOptions: IDropdownOption[] = [
  { key: 'Request', text: 'Request' },
  { key: 'Incident', text: 'Incident' },
  { key: 'Change', text: 'Change' },
];

const urgencyOptions: IDropdownOption[] = [
  { key: 'Alta', text: 'Alta' },
  { key: 'Media', text: 'Media' },
  { key: 'Baja', text: 'Baja' },
];

const securityOptions: IDropdownOption[] = [
  { key: 'No', text: 'No' },
  { key: 'Sí', text: 'Sí' },
];

const kbOptions: IDropdownOption[] = [
  { key: 'No', text: 'No' },
  { key: 'Sí', text: 'Sí' },
];

// ─── Step badge styles ──────────────────────────────────────────────────────────

function getStepStyle(resultado: string, isCurrent: boolean): React.CSSProperties {
  if (resultado === 'Aprobado') return { background: '#d4edda', border: '1px solid #28a745', borderRadius: 6, padding: '8px 12px', textAlign: 'center', fontSize: 13 };
  if (resultado === 'Rechazado') return { background: '#f8d7da', border: '1px solid #dc3545', borderRadius: 6, padding: '8px 12px', textAlign: 'center', fontSize: 13 };
  if (isCurrent) return { background: '#fff3cd', border: '1px solid #ffc107', borderRadius: 6, padding: '8px 12px', textAlign: 'center', fontSize: 13 };
  return { background: '#f8f9fa', border: '1px solid #dee2e6', borderRadius: 6, padding: '8px 12px', textAlign: 'center', fontSize: 13 };
}

// ─── Component ─────────────────────────────────────────────────────────────────

const DetalleTicket: React.FC<IDetalleTicketProps> = ({ isOpen, onDismiss, ticketId }) => {
  // ── Loading / feedback state ─────────────────────────────────────────────
  const [isLoading, setIsLoading] = React.useState(false);
  const [isSubmitting, setIsSubmitting] = React.useState(false);
  const [errorMessage, setErrorMessage] = React.useState('');
  const [successMessage, setSuccessMessage] = React.useState('');

  // ── Ticket data ──────────────────────────────────────────────────────────
  const [ticket, setTicket] = React.useState<ITicketData | null>(null);
  const [categorias, setCategorias] = React.useState<IDropdownOption[]>([]);
  const [adjuntos, setAdjuntos] = React.useState<IAdjunto[]>([]);
  const [aprobaciones, setAprobaciones] = React.useState<IAprobacionRow[]>([]);
  const [comments, setComments] = React.useState<IComment[]>([]);
  const [currentUserId, setCurrentUserId] = React.useState<number>(0);
  const dataLoaded = React.useRef(false);

  // ── Role flags ───────────────────────────────────────────────────────────
  const [iamApproval, setIamApproval] = React.useState(false);
  const [isProcessManager, setIsProcessManager] = React.useState(false);

  // ── Action panel visibility ──────────────────────────────────────────────
  const [showReasignarPanel, setShowReasignarPanel] = React.useState(false);
  const [showReaperturaPanel, setShowReaperturaPanel] = React.useState(false);
  const [showRecategorizarBar, setShowRecategorizarBar] = React.useState(false);
  const [categoryInputText, setCategoryInputText] = React.useState('');
  const [categoryMenuOpen, setCategoryMenuOpen] = React.useState(false);
  const categoryInputWrapperRef = React.useRef<HTMLDivElement>(null);

  // ── Form inputs (actions) ────────────────────────────────────────────────
  const [comentario, setComentario] = React.useState('');
  const [motivoReapertura, setMotivoReapertura] = React.useState('');
  const [solucionDetallada, setSolucionDetallada] = React.useState('');
  const [registrarKB, setRegistrarKB] = React.useState<string>('No');
  const [newCategoryId, setNewCategoryId] = React.useState<number | null>(null);
  const [selectedCommentFile, setSelectedCommentFile] = React.useState<File | null>(null);

  // ── Reassign people picker ───────────────────────────────────────────────
  const [reassignQuery, setReassignQuery] = React.useState('');
  const [reassignSuggestions, setReassignSuggestions] = React.useState<{ id: number; loginName: string; title: string }[]>([]);
  const [selectedReassign, setSelectedReassign] = React.useState<{ id: number; loginName: string; title: string } | null>(null);
  const commentFileRef = React.useRef<HTMLInputElement>(null);
  const [evaluacionRiesgo, setEvaluacionRiesgo] = React.useState<IEvaluacionRiesgoValues | null>(null);
  const [userGroups, setUserGroups] = React.useState<{ id: number; title: string }[]>([]);

  // ── Lifecycle ────────────────────────────────────────────────────────────

  React.useEffect(() => {
    if (isOpen && ticketId && !dataLoaded.current) {
      loadAll(ticketId).catch(console.error);
    }
    if (!isOpen) {
      resetState();
    }
  }, [isOpen, ticketId]);

  const resetState = (): void => {
    setTicket(null);
    setCategorias([]);
    setAdjuntos([]);
    setAprobaciones([]);
    setComments([]);
    setIamApproval(false);
    setIsProcessManager(false);
    setShowReasignarPanel(false);
    setShowReaperturaPanel(false);
    setShowRecategorizarBar(false);
    setCategoryInputText('');
    setCategoryMenuOpen(false);
    setComentario('');
    setMotivoReapertura('');
    setSolucionDetallada('');
    setRegistrarKB('No');
    setNewCategoryId(null);
    setSelectedCommentFile(null);
    setReassignQuery('');
    setReassignSuggestions([]);
    setSelectedReassign(null);
    setResolvedSteps([]);
    setEvaluacionRiesgo(null);
    setUserGroups([]);
    setErrorMessage('');
    setSuccessMessage('');
    dataLoaded.current = false;
  };

  // ── Data loading ─────────────────────────────────────────────────────────

  const loadAll = async (id: number): Promise<void> => {
    setIsLoading(true);
    try {
      const ticketItems: any[] = await ListSvc.getItems(
        'Tickets',
        undefined,
        `$filter=Id eq ${id}&$select=Id,Title,Status,TipoTicket,Descripcion,Department,Planta,Prioridad,TemplateConfiguracion,SolucionDetallada,Reaperturado,MotivoReapertura,Atencion,ImpactoSeguridadInformacion,DescripcionImpactoSeguridad,Created,Modified,CategoriaId,AuthorId,ProcessManagerId,ManagerId,ANombreDeId,Author/Title,Author/Id,Categoria/Title,ProcessManager/Title,ProcessManager/Id,Manager/Title,Manager/Id,Manager/EMail,ANombreDe/Title&$expand=Author,Categoria,ProcessManager,Manager,ANombreDe&$top=1`
      );
      const ticketItem = ticketItems?.[0];
      if (!ticketItem) throw new Error('Ticket no encontrado.');

      const currentUser = await UserSvc.GetCurrentUser();
      const uid: number = currentUser?.Id ?? 0;
      setCurrentUserId(uid);

      // Load Grupos where the current user is an Integrante
      const gruposUsuario: any[] = await ListSvc.getItems(
        'Grupos',
        undefined,
        `$select=Id,Title,Integrantes/Id&$expand=Integrantes&$filter=Integrantes/Id eq ${uid}&$top=500`
      ).catch(() => []);
      const mappedGroups = (gruposUsuario || []).map((g: any) => ({ id: g.Id, title: g.Title }));
      setUserGroups(mappedGroups);

      const tmplRaw: string | null = ticketItem?.TemplateConfiguracion ?? null;
      const tmpl: ITemplate | null = tmplRaw ? JSON.parse(tmplRaw) : null;

      const td: ITicketData = {
        id: ticketItem.Id,
        title: ticketItem.Title ?? '',
        solicitante: ticketItem.Author?.Title ?? '',
        solicitanteId: ticketItem.Author?.Id ?? 0,
        fechaRegistro: ticketItem.Created ? new Date(ticketItem.Created).toLocaleString('es-MX') : '',
        tipoTicket: ticketItem.TipoTicket ?? '',
        categoriaId: ticketItem.CategoriaId ?? null,
        categoriaTitle: ticketItem.Categoria?.Title ?? '',
        descripcion: ticketItem.Descripcion ?? '',
        departamento: ticketItem.Department ?? '',
        planta: ticketItem.Planta ?? '',
        aNombreDe: ticketItem.ANombreDe?.Title ?? null,
        prioridad: ticketItem.Prioridad ?? '',
        status: ticketItem.Status ?? '',
        templateConfiguracion: tmpl,
        solucionDetallada: ticketItem.SolucionDetallada ?? '',
        reaperturado: ticketItem.Reaperturado ?? false,
        motivoReapertura: ticketItem.MotivoReapertura ?? '',
        atencion: ticketItem.Atencion ?? false,
        impactoSeguridadInformacion: ticketItem.ImpactoSeguridadInformacion ?? 'No',
        descripcionImpactoSeguridad: ticketItem.DescripcionImpactoSeguridad ?? '',
        processManagerId: ticketItem.ProcessManagerId ?? 0,
        processManagerName: ticketItem.ProcessManager?.Title ?? '',
        managerId: ticketItem.ManagerId ?? null,
        managerName: ticketItem.Manager?.Title ?? null,
        managerEmail: ticketItem.Manager?.EMail ?? null,
        modified: ticketItem.Modified ?? '',
      };
      setTicket(td);
      setSolucionDetallada(td.solucionDetallada);

      const isPM = uid === td.processManagerId && td.status === 'Assigned';
      const isApproverStep = uid === td.managerId && td.status !== 'Assigned' && td.status !== 'Cerrado';
      setIsProcessManager(isPM);

      // Load entity type names for write operations
      const [ticketsType, aprobacionesType, comentariosType, knowledgeBaseType] = await Promise.all([
        ListSvc.getListItemEntityTypeFullName('Tickets'),
        ListSvc.getListItemEntityTypeFullName('Aprobaciones'),
        ListSvc.getListItemEntityTypeFullName('Comentarios').catch(() => 'SP.Data.ComentariosListItem'),
        ListSvc.getListItemEntityTypeFullName('Knowledge Base').catch(() => 'SP.Data.Knowledge_x0020_BaseListItem'),
      ]);
      setEntityTypes({ tickets: ticketsType, aprobaciones: aprobacionesType, comentarios: comentariosType, knowledgeBase: knowledgeBaseType });

      // Load categorias for recategorize dropdown
      const cats: any[] = await ListSvc.getItems(
        'Categorias',
        undefined,
        '$select=Id,Title,TipoCategoria,CategoriaPadre/Title,CategoriaPadreId&$expand=CategoriaPadre&$filter=CategoriaPadreId ne null&$top=300'
      );
      setCategorias((cats || []).map((c: any) => ({
        key: c.Id,
        text: `${c.CategoriaPadre?.Title ?? ''} - ${c.Title}`,
      })));

      // Load approval process
      const aprobItems: any[] = await ListSvc.getItems(
        'Aprobaciones',
        undefined,
        `$select=Id,Title,Resultado,GrupoId,Responsable/Id,Responsable/Title,Responsable/EMail,Modified&$expand=Responsable&$filter=TicketId eq ${id}&$orderby=Id asc`
      );
      const rows: IAprobacionRow[] = (aprobItems || []).map((a: any) => ({
        id: a.Id,
        title: a.Title ?? '',
        resultado: a.Resultado ?? 'Pendiente',
        responsable: a.Responsable?.Title ?? null,
        responsableEmail: a.Responsable?.EMail ?? null,
        modified: a.Modified ?? null,
        grupoId: a.GrupoId ?? null,
      }));
      setAprobaciones(rows);

      // Determine if current user is an approver
      const pendingRow = rows.find(r => r.resultado === 'Pendiente');
      if (pendingRow && td.status !== 'Assigned' && td.status !== 'Cerrado') {
        const isGroupApprover = pendingRow.grupoId != null
          ? mappedGroups.some(g => g.id === pendingRow.grupoId)
          : false;
        setIamApproval(isApproverStep || isGroupApprover || isPM);
      } else {
        setIamApproval(isPM);
      }

      // Resolve responsables for each approval step
      if (tmpl && tmpl.approvalPath.length > 0) {
        await resolveApprovalSteps(tmpl.approvalPath, rows, td.managerName, td.managerEmail, uid);
      }

      // Load adjuntos from Expediente folder
      try {
        const filesResp = await ListSvc.getAllFiles(
          'Expediente/' + id
        ).catch(() => null);
        if (filesResp) {
          setAdjuntos((filesResp.value || []).map((f: any) => ({ name: f.Name, url: f.ServerRelativeUrl })));
        }
      } catch {
        // Adjuntos may not be accessible, skip silently
      }

      // Load comments
      try {
        const commentItems: any[] = await ListSvc.getItems(
          'Comentarios',
          undefined,
          `$select=Id,Title,Author/Title,Created,Comentario&$expand=Author&$filter=TicketId eq ${id}&$orderby=Id asc`
        );
        const mappedComments = await Promise.all((commentItems || []).map(async (c: any) => {
          let adjuntos: { name: string; url: string }[] = [];
          try {
            const attachments = await ListSvc.getListItemAttachments('Comentarios', c.Id);
            adjuntos = (attachments || []).map((a: any) => ({ name: a.FileName, url: a.ServerRelativeUrl }));
          } catch { /* no attachments */ }
          return {
            id: c.Id,
            author: c.Author?.Title ?? '',
            text: c.Comentario ?? c.Title ?? '',
            date: c.Created ? new Date(c.Created).toLocaleString('es-MX') : '',
            adjuntos,
          };
        }));
        setComments(mappedComments);
      } catch {
        // Comments list may not exist
      }

      // Load EvaluacionesRiesgo for Change tickets
      if (td.tipoTicket === 'Change') {
        try {
          const evalItems: any[] = await ListSvc.getItems(
            'EvaluacionesRiesgo',
            undefined,
            `$select=TipoCambio,SistemasInvolucrados,FechaPropuesta,MotivoCambio,EvaluacionRiesgos,InterrupcionServicio,DuracionEstimada,UsuariosAfectados,ProcesosCriticosImpactados,PasosImplementacion,PuedoRevertir,ProcedimientoRollback,CriteriosExito&$filter=TicketId eq ${id}&$top=1`
          );
          const ev = evalItems?.[0];
          if (ev) {
            const fechaDate = ev.FechaPropuesta ? new Date(ev.FechaPropuesta) : null;
            setEvaluacionRiesgo({
              tipoCambio: ev.TipoCambio ?? '',
              sistemasInvolucrados: ev.SistemasInvolucrados ?? '',
              fechaPropuesta: fechaDate,
              horaPropuesta: fechaDate,
              motivoCambio: ev.MotivoCambio ?? '',
              evaluacionRiesgos: ev.EvaluacionRiesgos ?? '',
              interrupcionServicio: ev.InterrupcionServicio ?? 'No',
              duracionEstimada: ev.DuracionEstimada != null ? String(ev.DuracionEstimada) : '',
              usuariosAfectados: ev.UsuariosAfectados != null ? String(ev.UsuariosAfectados) : '',
              procesosCriticosImpactados: ev.ProcesosCriticosImpactados ?? '',
              pasosImplementacion: ev.PasosImplementacion ?? '',
              puedeRevertir: ev.PuedoRevertir ?? 'Sí',
              procedimientoRollback: ev.ProcedimientoRollback ?? '',
              criteriosExito: ev.CriteriosExito ?? '',
            });
          }
        } catch {
          // EvaluacionesRiesgo may not exist for this ticket
        }
      }

      dataLoaded.current = true;
    } catch (e) {
      console.error('Error loading ticket detail:', e);
      setErrorMessage('Error al cargar el detalle del ticket.');
    } finally {
      setIsLoading(false);
    }
  };

  const [entityTypes, setEntityTypes] = React.useState<{ tickets: string; aprobaciones: string; comentarios: string; knowledgeBase: string }>({ tickets: '', aprobaciones: '', comentarios: '', knowledgeBase: '' });

  // ── Resolved approval steps ───────────────────────────────────────────────
  const [resolvedSteps, setResolvedSteps] = React.useState<IResolvedStep[]>([]);

  const resolveApprovalSteps = async (
    approvalPath: IApprovalStep[],
    rows: IAprobacionRow[],
    managerName: string | null,
    managerEmail: string | null,
    currentUid: number
  ): Promise<void> => {
    const resolved: IResolvedStep[] = [];
    let userIsApprover = false;

    for (let i = 0; i < approvalPath.length; i++) {
      const ap = approvalPath[i];
      const row = rows[i];

      // If the aprobacion row already has an assigned Responsable, use it
      if (row?.responsable) {
        resolved.push({
          stepName: ap.stepName,
          responsable: row.responsable,
          responsableEmail: row.responsableEmail,
        });
        continue;
      }

      // useUserDepartment → responsable is the ticket Manager
      if (ap.useUserDepartment) {
        if (currentUid === rows[i]?.id) userIsApprover = true; // handled by outer logic
        resolved.push({
          stepName: ap.stepName,
          responsable: managerName,
          responsableEmail: managerEmail,
        });
        continue;
      }

      // Group-based step → query Grupos list by Title = stepName
      try {
        const groupItems: any[] = await ListSvc.getItems(
          'Grupos',
          undefined,
          `$filter=Title eq '${ap.stepName.replace(/'/g, "''")}'&$select=Id,Title,Integrantes/Id,Integrantes/Title,Integrantes/EMail&$expand=Integrantes&$top=1`
        );
        const grupo = groupItems?.[0];
        const integrantes: any[] = grupo?.Integrantes ?? [];
        const primero = integrantes[0];
        resolved.push({
          stepName: ap.stepName,
          responsable: primero?.Title ?? null,
          responsableEmail: primero?.EMail ?? null,
        });
      } catch {
        resolved.push({ stepName: ap.stepName, responsable: null, responsableEmail: null });
      }
    }

    setResolvedSteps(resolved);
  };

  // ── Reassign search ───────────────────────────────────────────────────────

  const onReassignSearch = async (value: string): Promise<void> => {
    setReassignQuery(value);
    if (value.length < 3) { setReassignSuggestions([]); return; }
    try {
      const results = await UserSvc.SearchUsers(value);
      setReassignSuggestions((results || []).slice(0, 8).map((u: any) => ({
        id: u.EntityData?.SPUserID ?? 0,
        loginName: u.Key ?? '',
        title: u.DisplayText ?? u.Description ?? '',
      })));
    } catch { setReassignSuggestions([]); }
  };

  // ── Actions ───────────────────────────────────────────────────────────────

  const handlePublishComment = async (): Promise<void> => {
    if (!ticket) return;
    if (!comentario.trim() && !selectedCommentFile) {
      setErrorMessage('Escribe un comentario o selecciona un archivo para publicar.');
      return;
    }

    setErrorMessage('');
    setIsSubmitting(true);
    try {
      const newComment = await ListSvc.postListItem(
        'Comentarios',
        JSON.stringify({
          '__metadata': { type: entityTypes.comentarios },
          Title: ticket.status,
          Comentario: comentario,
          TicketId: ticket.id,
          ResponsableId: currentUserId,
        })
      );

      if (selectedCommentFile) {
        const comentarioId: number = newComment?.d?.Id ?? newComment?.Id;
        if (comentarioId) {
          const arrayBuffer = await new Promise<ArrayBuffer>((resolve, reject) => {
            const reader = new FileReader();
            reader.onload = (e) => resolve(e.target!.result as ArrayBuffer);
            reader.onerror = (e) => reject(e);
            reader.readAsArrayBuffer(selectedCommentFile);
          });
          await ListSvc.postListItemAttachment('Comentarios', comentarioId, selectedCommentFile.name, arrayBuffer);
        }
      }

      setSuccessMessage('Comentario publicado correctamente.');
      setComentario('');
      setSelectedCommentFile(null);
      dataLoaded.current = false;
      await loadAll(ticket.id);
    } catch (e) {
      console.error(e);
      setErrorMessage('Error al publicar el comentario.');
    } finally {
      setIsSubmitting(false);
    }
  };

  const handleApprove = async (): Promise<void> => {
    if (!ticket) return;
    setErrorMessage('');
    setSubmitAction('approve');
    setIsSubmitting(true);
    try {
      // 1. Mark current Aprobaciones row as Aprobado and assign the current user as Responsable
      const pendingRow = aprobaciones.find(r => r.resultado === 'Pendiente');
      if (pendingRow) {
        await ListSvc.putListItem(
          'Aprobaciones',
          pendingRow.id,
          JSON.stringify({
            '__metadata': { type: entityTypes.aprobaciones },
            Resultado: 'Aprobado',
            ResponsableId: currentUserId,
          })
        );
      }

      // 2. Locate current step in the approvalPath and find the next one
      const tmpl = ticket.templateConfiguracion;
      const ap = tmpl?.approvalPath ?? [];
      const currentIdx = ap.findIndex(s => s.stepName === ticket.status);
      const nextStep = ap[currentIdx + 1] ?? null;

      // 3. Determine new ticket status
      //    - There is a next approval step  → use nextStep.stepName
      //    - No more approval steps         → 'Assigned' (moves to Process Manager)
      const newStatus = nextStep ? nextStep.stepName : 'Assigned';
      await ListSvc.putListItem(
        'Tickets',
        ticket.id,
        JSON.stringify({ '__metadata': { type: entityTypes.tickets }, Status: newStatus })
      );

      // 4. Create the next Aprobaciones record
      //    nextStep != null → next approval step (group or manager)
      //    nextStep == null → PM step (Assigned)
      const newAprobacionBody: Record<string, unknown> = {
        '__metadata': { type: entityTypes.aprobaciones },
        Title: newStatus,
        Resultado: 'Pendiente',
        TicketId: ticket.id,
      };

      if (nextStep) {
        // approvalGroup is a number → assign to that group
        if (nextStep.approvalGroup != null) {
          newAprobacionBody['GrupoId'] = nextStep.approvalGroup;
        } else {
          // useUserDepartment → responsable is the ticket Manager
          newAprobacionBody['ResponsableId'] = ticket.managerId ?? currentUserId;
        }
      } else {
        // No more approval steps → assign to Process Manager
        newAprobacionBody['ResponsableId'] = ticket.processManagerId;
      }

      await ListSvc.postListItem('Aprobaciones', JSON.stringify(newAprobacionBody));

      // 5. Optional comment
      if (comentario.trim()) {
        const newComment = await ListSvc.postListItem(
          'Comentarios',
          JSON.stringify({
            '__metadata': { type: entityTypes.comentarios },
            Title: ticket.status,
            Comentario: comentario,
            TicketId: ticket.id,
            ResponsableId: currentUserId,
          })
        );
        if (selectedCommentFile) {
          const comentarioId: number = newComment?.d?.Id ?? newComment?.Id;
          if (comentarioId) {
            const arrayBuffer = await new Promise<ArrayBuffer>((resolve, reject) => {
              const reader = new FileReader();
              reader.onload = (e) => resolve(e.target!.result as ArrayBuffer);
              reader.onerror = (e) => reject(e);
              reader.readAsArrayBuffer(selectedCommentFile);
            });
            await ListSvc.postListItemAttachment('Comentarios', comentarioId, selectedCommentFile.name, arrayBuffer);
          }
        }
      }

      setSuccessMessage('Ticket aprobado correctamente.');
      dataLoaded.current = false;
      await loadAll(ticket.id);
    } catch (e) {
      console.error(e);
      setErrorMessage('Error al aprobar el ticket.');
    } finally {
      setIsSubmitting(false);
    }
  };

  const handleReject = async (): Promise<void> => {
    if (!ticket) return;
    if (!comentario.trim()) { setErrorMessage('Ingrese un comentario para rechazar el ticket.'); return; }
    setErrorMessage('');
    setIsSubmitting(true);
    try {
      const pendingRow = aprobaciones.find(r => r.resultado === 'Pendiente');
      if (pendingRow) {
        await ListSvc.putListItem('Aprobaciones', pendingRow.id, JSON.stringify({ '__metadata': { type: entityTypes.aprobaciones }, Resultado: 'Rechazado' }));
      }
      await ListSvc.putListItem('Tickets', ticket.id, JSON.stringify({ '__metadata': { type: entityTypes.tickets }, Status: 'Rechazado' }));
      await ListSvc.postListItem('Comentarios', JSON.stringify({ '__metadata': { type: entityTypes.comentarios }, Title: comentario, TicketId: ticket.id }));
      setSuccessMessage('Ticket rechazado.');
      dataLoaded.current = false;
      await loadAll(ticket.id);
    } catch (e) {
      console.error(e);
      setErrorMessage('Error al rechazar el ticket.');
    } finally {
      setIsSubmitting(false);
    }
  };

  const handleClose = async (): Promise<void> => {
    if (!ticket) return;
    setErrorMessage('');
    setIsSubmitting(true);
    try {
      let selectedFileBuffer: ArrayBuffer | null = null;
      if (selectedCommentFile) {
        selectedFileBuffer = await new Promise<ArrayBuffer>((resolve, reject) => {
          const reader = new FileReader();
          reader.onload = (e) => resolve(e.target!.result as ArrayBuffer);
          reader.onerror = (e) => reject(e);
          reader.readAsArrayBuffer(selectedCommentFile);
        });
      }

      const updatePayload: Record<string, unknown> = { '__metadata': { type: entityTypes.tickets }, Status: 'Cerrado' };
      if (registrarKB === 'Sí' && solucionDetallada.trim()) {
        updatePayload.SolucionDetallada = solucionDetallada;
      }
      await ListSvc.putListItem('Tickets', ticket.id, JSON.stringify(updatePayload));
      if (registrarKB === 'Sí') {
        const solucionParaKB = (solucionDetallada.trim() || ticket.solucionDetallada || '').trim();
        const kbItem = await ListSvc.postListItem(
          'Knowledge Base',
          JSON.stringify({
            '__metadata': { type: entityTypes.knowledgeBase },
            TicketOrigenId: ticket.id,
            Title: ticket.title,
            Categoria: ticket.categoriaTitle,
            SolucionDetallada: solucionParaKB,
            DescripcionDetallada: ticket.descripcion,
            Departamento: ticket.departamento,
          })
        );
        if (selectedCommentFile && selectedFileBuffer) {
          const kbItemId: number = kbItem?.d?.Id ?? kbItem?.Id;
          if (kbItemId) {
            await ListSvc.postListItemAttachment('Knowledge Base', kbItemId, selectedCommentFile.name, selectedFileBuffer);
          }
        }
      }
      if (comentario.trim() || selectedCommentFile) {
        const closeComment = await ListSvc.postListItem(
          'Comentarios',
          JSON.stringify({ '__metadata': { type: entityTypes.comentarios }, Title: comentario || 'Cierre de ticket', TicketId: ticket.id, Comentario: comentario })
        );
        if (selectedCommentFile && selectedFileBuffer) {
          const closeCommentId: number = closeComment?.d?.Id ?? closeComment?.Id;
          if (closeCommentId) {
            await ListSvc.postListItemAttachment('Comentarios', closeCommentId, selectedCommentFile.name, selectedFileBuffer);
          }
        }
      }
      setSuccessMessage('Ticket cerrado correctamente.');
      dataLoaded.current = false;
      await loadAll(ticket.id);
    } catch (e) {
      console.error(e);
      setErrorMessage('Error al cerrar el ticket.');
    } finally {
      setIsSubmitting(false);
    }
  };

  const handleRecategorizar = async (): Promise<void> => {
    if (!ticket || !newCategoryId) { setErrorMessage('Seleccione una categoría.'); return; }
    setErrorMessage('');
    setIsSubmitting(true);
    try {
      // 1. Load category to get its TemplateAtencion lookup
      const catItems: any[] = await ListSvc.getItems(
        'Categorias',
        undefined,
        `$select=Id,TemplateAtencionId&$filter=Id eq ${newCategoryId}&$top=1`
      );
      const cat = catItems?.[0];
      const templateId: number | null = cat?.TemplateAtencionId ?? null;

      let newTemplateConfig: ITemplate | null = null;
      let newStatus = 'Assigned';

      if (templateId) {
        // 2. Load template to get RutaAprobacionEscalacion
        const tmplItem = await ListSvc.getItemById('Templates', templateId);
        const rutaJson: string | null = tmplItem?.RutaAprobacionEscalacion ?? null;

        if (rutaJson) {
          newTemplateConfig = JSON.parse(rutaJson) as ITemplate;
          // 3. Determine new status from approval path
          newStatus = newTemplateConfig.approvalPath?.length > 0
            ? newTemplateConfig.approvalPath[0].stepName
            : 'Assigned';
        }
      }

      // 4. Update ticket: CategoriaId, TemplateConfiguracion, Status
      const ticketPayload: Record<string, unknown> = {
        '__metadata': { type: entityTypes.tickets },
        CategoriaId: newCategoryId,
        Status: newStatus,
      };
      if (newTemplateConfig) {
        ticketPayload['TemplateConfiguracion'] = JSON.stringify(newTemplateConfig);
      }
      await ListSvc.putListItem('Tickets', ticket.id, JSON.stringify(ticketPayload));

      // 5. Delete existing Aprobaciones and create the first one for the new path
      const existingAprobaciones: any[] = await ListSvc.getItems(
        'Aprobaciones',
        undefined,
        `$select=Id&$filter=TicketId eq ${ticket.id}`
      );
      await Promise.all((existingAprobaciones || []).map((a: any) =>
        ListSvc.deleteListItem('Aprobaciones', a.Id)
      ));

      // 6. Create first Aprobaciones record for the new path
      const newAprobBody: Record<string, unknown> = {
        '__metadata': { type: entityTypes.aprobaciones },
        Title: newStatus,
        Resultado: 'Pendiente',
        TicketId: ticket.id,
      };

      if (newTemplateConfig && newTemplateConfig.approvalPath?.length > 0) {
        const firstStep = newTemplateConfig.approvalPath[0];
        if (firstStep.approvalGroup != null) {
          newAprobBody['GrupoId'] = firstStep.approvalGroup;
        } else if (firstStep.useUserDepartment) {
          newAprobBody['ResponsableId'] = ticket.managerId ?? ticket.processManagerId;
        } else {
          newAprobBody['ResponsableId'] = ticket.processManagerId;
        }
      } else {
        newAprobBody['ResponsableId'] = ticket.processManagerId;
      }
      await ListSvc.postListItem('Aprobaciones', JSON.stringify(newAprobBody));

      // 7. Post default recategorization comment
      const newCategoryText = categorias.find(c => c.key === newCategoryId)?.text ?? String(newCategoryId);
      const comentarioRecat = `Ticket recategorizado. Categoría original: "${ticket.categoriaTitle}" → Nueva categoría: "${newCategoryText}".`;
      await ListSvc.postListItem(
        'Comentarios',
        JSON.stringify({
          '__metadata': { type: entityTypes.comentarios },
          Title: ticket.status,
          Comentario: comentarioRecat,
          TicketId: ticket.id,
          ResponsableId: currentUserId,
        })
      );

      setSuccessMessage('Ticket recategorizado correctamente.');
      setShowRecategorizarBar(false);
      dataLoaded.current = false;
      await loadAll(ticket.id);
    } catch (e) {
      console.error(e);
      setErrorMessage('Error al recategorizar el ticket.');
    } finally {
      setIsSubmitting(false);
    }
  };

  const handleReaperturar = async (): Promise<void> => {
    if (!ticket) return;
    if (!motivoReapertura.trim()) { setErrorMessage('Ingrese el motivo de la reapertura.'); return; }
    setErrorMessage('');
    setIsSubmitting(true);
    try {
      await ListSvc.putListItem('Tickets', ticket.id, JSON.stringify({
        '__metadata': { type: entityTypes.tickets },
        Status: 'Assigned',
        Reaperturado: true,
        MotivoReapertura: motivoReapertura,
      }));
      setSuccessMessage('Ticket reaperturado correctamente.');
      setShowReaperturaPanel(false);
      dataLoaded.current = false;
      await loadAll(ticket.id);
    } catch (e) {
      console.error(e);
      setErrorMessage('Error al reaperturar el ticket.');
    } finally {
      setIsSubmitting(false);
    }
  };

  const handleConfirmReassign = async (): Promise<void> => {
    if (!ticket || !selectedReassign) { setErrorMessage('Seleccione un usuario para reasignar.'); return; }
    setErrorMessage('');
    setIsSubmitting(true);
    try {
      const ensured = await UserSvc.EnsureUser(selectedReassign.loginName);
      const newId: number = ensured?.Id ?? (ensured as any)?.d?.Id;
      if (!newId) { setErrorMessage('No se pudo resolver el usuario.'); return; }
      await ListSvc.putListItem('Tickets', ticket.id, JSON.stringify({ '__metadata': { type: entityTypes.tickets }, ProcessManagerId: newId }));
      setSuccessMessage('Ticket reasignado correctamente.');
      setShowReasignarPanel(false);
      setSelectedReassign(null);
      setReassignQuery('');
      dataLoaded.current = false;
      await loadAll(ticket.id);
    } catch (e) {
      console.error(e);
      setErrorMessage('Error al reasignar el ticket.');
    } finally {
      setIsSubmitting(false);
    }
  };

  const handleSetAtencion = async (): Promise<void> => {
    if (!ticket) return;
    setIsSubmitting(true);
    try {
      await ListSvc.putListItem('Tickets', ticket.id, JSON.stringify({ '__metadata': { type: entityTypes.tickets }, Atencion: true }));
      setSuccessMessage('Ticket marcado en atención.');
      dataLoaded.current = false;
      await loadAll(ticket.id);
    } catch (e) {
      console.error(e);
      setErrorMessage('Error al actualizar el ticket.');
    } finally {
      setIsSubmitting(false);
    }
  };
  
  // Placeholder needed to satisfy compile
  const [submitAction, setSubmitAction] = React.useState<string>('');

  // ── Status badge helper ───────────────────────────────────────────────────

  const isClosed = ticket?.status === 'Cerrado';
  const canReaperturar = isClosed && ticket
    ? (new Date().getTime() - new Date(ticket.modified).getTime()) / (1000 * 60 * 60 * 24) >= 30
    : false;

  // ── Render helpers ────────────────────────────────────────────────────────

  const renderApprovalPath = (): JSX.Element => {
    if (!ticket) return <></>;
    const tmpl = ticket.templateConfiguracion;
    const steps: { stepName: string; resultado: string; responsable: string | null; responsableEmail: string | null; date: string | null }[] = [];

    if (tmpl) {
      tmpl.approvalPath.forEach((ap, i) => {
        const row = aprobaciones[i];
        const resolved = resolvedSteps[i];
        steps.push({
          stepName: ap.stepName,
          resultado: row?.resultado ?? 'Pendiente',
          responsable: resolved?.responsable ?? null,
          responsableEmail: row?.resultado === 'Pendiente' ? (resolved?.responsableEmail ?? null) : null,
          date: row?.resultado !== 'Pendiente' && row?.modified ? new Date(row.modified).toLocaleString('es-MX') : null,
        });
      });
    }

    // Process manager step
    const pmStyle: React.CSSProperties = ticket.status === 'Assigned'
      ? (ticket.reaperturado
        ? { background: '#cce5ff', border: '1px solid #004085', borderRadius: 6, padding: '8px 12px', textAlign: 'center', fontSize: 13 }
        : { background: '#fff3cd', border: '1px solid #ffc107', borderRadius: 6, padding: '8px 12px', textAlign: 'center', fontSize: 13 })
      : ticket.status === 'Cerrado'
        ? { background: '#d4edda', border: '1px solid #28a745', borderRadius: 6, padding: '8px 12px', textAlign: 'center', fontSize: 13 }
        : { background: '#f8f9fa', border: '1px solid #dee2e6', borderRadius: 6, padding: '8px 12px', textAlign: 'center', fontSize: 13 };

    return (
      <div style={{ display: 'flex', flexWrap: 'wrap', gap: 8, marginBottom: 12 }}>
        {steps.map((s, i) => (
          <div key={i} style={{ ...getStepStyle(s.resultado, ticket.status === s.stepName), minWidth: 140 }}>
            <strong style={{ display: 'block', fontSize: 12 }}>{s.stepName}</strong>
            {s.responsableEmail && s.resultado === 'Pendiente'
              ? <a href={`https://teams.microsoft.com/l/chat/0/0?users=${s.responsableEmail}`} target="_blank" rel="noopener noreferrer" style={{ fontSize: 12 }}>{s.responsable}</a>
              : <span style={{ fontSize: 12 }}>{s.responsable}</span>}
            {s.date && <span style={{ display: 'block', fontSize: 11, color: '#6c757d' }}>{s.date}</span>}
          </div>
        ))}
        <div style={{ ...pmStyle, minWidth: 140 }}>
          <strong style={{ display: 'block', fontSize: 12 }}>Ticket service</strong>
          <span style={{ fontSize: 12 }}>{ticket.processManagerName}</span>
          {isClosed && ticket.modified && (
            <span style={{ display: 'block', fontSize: 11, color: '#6c757d' }}>{new Date(ticket.modified).toLocaleString('es-MX')}</span>
          )}
        </div>
      </div>
    );
  };

  const renderTemplateFields = (): JSX.Element => {
    const ft = ticket?.templateConfiguracion?.formTemplate;
    if (!ft || ft.length === 0) return <></>;
    return (
      <>
        <Separator />
        {ft.map((field, i) => {
          switch (field.tag) {
            case 'h3': return <h4 key={i} style={{ margin: '8px 0 4px' }}>{field.text}</h4>;
            case 'p': return <p key={i} style={{ margin: '4px 0' }}>{field.text}</p>;
            case 'br': return <br key={i} />;
            case 'checkbox':
              return (
                <div key={i} style={{ display: 'flex', alignItems: 'center', gap: 8, margin: '4px 0' }}>
                  <input type="checkbox" disabled checked={!!field.checked} id={field.id} />
                  <label htmlFor={field.id} style={{ fontSize: 14, margin: 0 }}>{field.text}</label>
                </div>
              );
            case 'textarea':
              return (
                <TextField
                  key={i}
                  label={field.text}
                  multiline
                  rows={3}
                  disabled
                  value={field.value ?? ''}
                />
              );
            default: return null;
          }
        })}
        <Separator />
      </>
    );
  };

  // ── Main render ───────────────────────────────────────────────────────────

  return (
    <Panel
      isOpen={isOpen}
      onDismiss={onDismiss}
      headerText="Seguimiento al ticket"
      closeButtonAriaLabel="Cerrar"
      type={PanelType.large}
    >
      {isLoading && (
        <Stack horizontalAlign="center" styles={{ root: { marginTop: 40 } }}>
          <Spinner size={SpinnerSize.large} label="Cargando ticket..." />
        </Stack>
      )}

      {!isLoading && !ticketId && (
        <MessageBar messageBarType={MessageBarType.info}>
          No se ha seleccionado ningún ticket.
        </MessageBar>
      )}

      {!isLoading && ticket && (
        <Stack tokens={stackTokens}>
          {/* Feedback messages */}
          {errorMessage && (
            <MessageBar messageBarType={MessageBarType.error} onDismiss={() => setErrorMessage('')}>
              {errorMessage}
            </MessageBar>
          )}
          {successMessage && (
            <MessageBar messageBarType={MessageBarType.success} onDismiss={() => setSuccessMessage('')}>
              {successMessage}
            </MessageBar>
          )}

          {/* Approval status path */}
          <div style={{ border: '1px solid #eee', padding: 12, borderRadius: 6, boxShadow: '0 2px 8px rgba(0,0,0,0.12)' }}>
            <p style={{ fontSize: 13, color: '#6c757d', marginBottom: 8 }}>
              Presiona el nombre de la persona para enviarle un recordatorio de su revisión por Teams.
            </p>
            {renderApprovalPath()}
          </div>

          {/* Option bar */}
          {(isProcessManager || iamApproval) && (
            <div style={{ display: 'flex', gap: 8, justifyContent: 'flex-end', flexWrap: 'wrap' }}>
              {isProcessManager && (
                <TooltipHost content="Reasignar Ticket">
                  <DefaultButton iconProps={{ iconName: 'People' }} text="Reasignar" onClick={() => { setShowReasignarPanel(v => !v); setShowReaperturaPanel(false); }} />
                </TooltipHost>
              )}
              {isProcessManager && (
                <TooltipHost content="Recategorizar Ticket">
                  <DefaultButton iconProps={{ iconName: 'Tag' }} text="Recategorizar" onClick={() => {
                    const opening = !showRecategorizarBar;
                    setShowRecategorizarBar(opening);
                    if (opening) {
                      setCategoryInputText(ticket.categoriaTitle);
                      setCategoryMenuOpen(false);
                      setNewCategoryId(ticket.categoriaId);
                    } else {
                      setCategoryInputText('');
                      setCategoryMenuOpen(false);
                    }
                  }} />
                </TooltipHost>
              )}
              {isProcessManager && canReaperturar && (
                <TooltipHost content="Reaperturar Ticket">
                  <DefaultButton iconProps={{ iconName: 'Refresh' }} text="Reaperturar" onClick={() => { setShowReaperturaPanel(v => !v); setShowReasignarPanel(false); }} />
                </TooltipHost>
              )}
              {isProcessManager && !ticket.atencion && (
                <TooltipHost content="Marcar en Atención">
                  <DefaultButton iconProps={{ iconName: 'Clock' }} text="En Atención" onClick={handleSetAtencion} disabled={isSubmitting} />
                </TooltipHost>
              )}
            </div>
          )}

          {/* Reassign panel */}
          {showReasignarPanel && (
            <div style={{ border: '1px solid #eee', padding: 15, borderRadius: 6, boxShadow: '0 3px 10px rgba(0,0,0,0.2)' }}>
              <Label style={{ fontWeight: 600, marginBottom: 6 }}>Reasignar responsable</Label>
              <TextField
                placeholder="Buscar persona..."
                value={reassignQuery}
                onChange={(_, v) => onReassignSearch(v ?? '')}
              />
              {reassignSuggestions.length > 0 && (
                <ul style={{ listStyle: 'none', margin: '4px 0 0', padding: 0, border: '1px solid #ddd', borderRadius: 4, maxHeight: 180, overflowY: 'auto' }}>
                  {reassignSuggestions.map(s => (
                    <li
                      key={s.loginName}
                      style={{ padding: '6px 10px', cursor: 'pointer', background: selectedReassign?.loginName === s.loginName ? '#e9ecef' : 'white', fontSize: 13 }}
                      onClick={() => { setSelectedReassign(s); setReassignQuery(s.title); setReassignSuggestions([]); }}
                    >
                      {s.title}
                    </li>
                  ))}
                </ul>
              )}
              <div style={{ marginTop: 10, padding: 10, background: '#f8f9fa', borderLeft: '4px solid #ffc107', borderRadius: 4, fontSize: 13, color: '#6c757d' }}>
                Una vez que se reasigne a otro usuario, perderá el privilegio de aprobar o rechazar el ticket.
              </div>
              <Stack horizontal tokens={{ childrenGap: 8 }} styles={{ root: { marginTop: 12, justifyContent: 'flex-end' } }}>
                <PrimaryButton text="Reasignar" onClick={handleConfirmReassign} disabled={isSubmitting || !selectedReassign} />
                <DefaultButton text="Cancelar" onClick={() => { setShowReasignarPanel(false); setSelectedReassign(null); setReassignQuery(''); }} />
              </Stack>
            </div>
          )}

          {/* Reapertura panel */}
          {showReaperturaPanel && (
            <div style={{ border: '1px solid #eee', padding: 15, borderRadius: 6, boxShadow: '0 3px 10px rgba(0,0,0,0.2)' }}>
              <Label style={{ fontWeight: 600, marginBottom: 6 }}>Reaperturar Ticket</Label>
              <TextField
                label="Motivo de la reapertura"
                multiline
                rows={4}
                placeholder="Ingrese el motivo de la reapertura"
                value={motivoReapertura}
                onChange={(_, v) => setMotivoReapertura(v ?? '')}
              />
              <Stack horizontal tokens={{ childrenGap: 8 }} styles={{ root: { marginTop: 12, justifyContent: 'flex-end' } }}>
                <PrimaryButton text="Reaperturar" onClick={handleReaperturar} disabled={isSubmitting} />
                <DefaultButton text="Cancelar" onClick={() => setShowReaperturaPanel(false)} />
              </Stack>
            </div>
          )}

          {/* Main form fields */}
          <Stack horizontal tokens={{ childrenGap: 12 }} styles={{ root: { flexWrap: 'wrap' } }}>
            <Stack.Item grow={3} styles={{ root: { minWidth: 200 } }}>
              <TextField label="Solicitante" value={ticket.solicitante} readOnly />
            </Stack.Item>
            <Stack.Item grow={1} styles={{ root: { minWidth: 100 } }}>
              <TextField label="Folio" value={String(ticket.id)} readOnly />
            </Stack.Item>
            <Stack.Item grow={2} styles={{ root: { minWidth: 150 } }}>
              <TextField label="Fecha de registro" value={ticket.fechaRegistro} readOnly />
            </Stack.Item>
          </Stack>

          <Stack horizontal tokens={{ childrenGap: 12 }} styles={{ root: { flexWrap: 'wrap' } }}>
            <Stack.Item grow={1} styles={{ root: { minWidth: 150 } }}>
              <Dropdown
                label="Tipo"
                options={typeOptions}
                selectedKey={ticket.tipoTicket}
                disabled
              />
            </Stack.Item>
            <Stack.Item grow={3} styles={{ root: { minWidth: 200 } }}>
              {showRecategorizarBar ? (
                <div ref={categoryInputWrapperRef}>
                  <TextField
                    label="Categoría"
                    placeholder="Buscar categoría..."
                    value={categoryInputText}
                    onChange={(_, v) => {
                      setCategoryInputText(v ?? '');
                      setNewCategoryId(null);
                      setCategoryMenuOpen(true);
                    }}
                    onFocus={() => setCategoryMenuOpen(true)}
                    autoComplete="off"
                  />
                  {categoryMenuOpen && categoryInputWrapperRef.current && (() => {
                    const filtered = categorias.filter(c =>
                      !categoryInputText || (c.text ?? '').toLowerCase().includes(categoryInputText.toLowerCase())
                    );
                    return filtered.length > 0 ? (
                      <Callout
                        target={categoryInputWrapperRef.current}
                        onDismiss={() => setCategoryMenuOpen(false)}
                        isBeakVisible={false}
                        directionalHint={DirectionalHint.bottomLeftEdge}
                        calloutMaxHeight={280}
                        styles={{ calloutMain: { overflowY: 'auto' } }}
                        setInitialFocus={false}
                      >
                        {filtered.map(c => (
                          <div
                            key={c.key}
                            onMouseDown={(e) => {
                              e.preventDefault();
                              setCategoryInputText(c.text ?? '');
                              setNewCategoryId(c.key as number);
                              setCategoryMenuOpen(false);
                            }}
                            style={{ padding: '8px 16px', cursor: 'pointer', whiteSpace: 'nowrap', fontSize: 14 }}
                            onMouseEnter={(e) => { (e.currentTarget as HTMLDivElement).style.background = '#f3f2f1'; }}
                            onMouseLeave={(e) => { (e.currentTarget as HTMLDivElement).style.background = ''; }}
                          >
                            {c.text}
                          </div>
                        ))}
                      </Callout>
                    ) : null;
                  })()}
                </div>
              ) : (
                <TextField
                  label="Categoría"
                  value={ticket.categoriaTitle}
                  readOnly
                />
              )}
            </Stack.Item>
          </Stack>

          {/* Recategorizar action bar */}
          {showRecategorizarBar && (
            <Stack horizontal tokens={{ childrenGap: 8 }} styles={{ root: { justifyContent: 'flex-end' } }}>
              <PrimaryButton text="Recategorizar" onClick={handleRecategorizar} disabled={isSubmitting} />
              <DefaultButton text="Cancelar" onClick={() => { setShowRecategorizarBar(false); setNewCategoryId(null); setCategoryInputText(''); setCategoryMenuOpen(false); }} />
            </Stack>
          )}

          <TextField label="Título" value={ticket.title} readOnly />

          <TextField label="Descripción detallada" value={ticket.descripcion} multiline rows={4} readOnly />

          <Stack horizontal tokens={{ childrenGap: 12 }} styles={{ root: { flexWrap: 'wrap' } }}>
            <Stack.Item grow={1} styles={{ root: { minWidth: 140 } }}>
              <TextField label="Departamento" value={ticket.departamento} readOnly />
            </Stack.Item>
            <Stack.Item grow={1} styles={{ root: { minWidth: 140 } }}>
              <TextField label="Planta" value={ticket.planta} readOnly />
            </Stack.Item>
            {ticket.aNombreDe && (
              <Stack.Item grow={1} styles={{ root: { minWidth: 140 } }}>
                <TextField label="A nombre de" value={ticket.aNombreDe} readOnly />
              </Stack.Item>
            )}
          </Stack>

          {/* Dynamic template fields */}
          {renderTemplateFields()}

          <Dropdown
            label="Urgencia"
            options={urgencyOptions}
            selectedKey={ticket.prioridad}
            disabled
            styles={{ root: { maxWidth: 200 } }}
          />

          {/* Change validation */}
          {ticket.tipoTicket === 'Change' && (
            <Stack tokens={stackTokens}>
              <Dropdown
                label="¿Este cambio impacta la seguridad de la información?"
                options={securityOptions}
                selectedKey={ticket.impactoSeguridadInformacion}
                disabled
                styles={{ root: { maxWidth: 300 } }}
              />
              {ticket.impactoSeguridadInformacion === 'Sí' && (
                <TextField
                  label="Describe los impactos a la seguridad de la información"
                  multiline
                  rows={3}
                  value={ticket.descripcionImpactoSeguridad}
                  readOnly
                />
              )}
            </Stack>
          )}

          {/* Evaluación de Riesgo (Change tickets) */}
          {ticket.tipoTicket === 'Change' && evaluacionRiesgo && (
            <EvaluacionRiesgo readOnly initialValues={evaluacionRiesgo} />
          )}

          <Separator />

          {/* Attached files */}
          {adjuntos.length > 0 && (
            <div>
              <Label>Archivos adjuntos</Label>
              <ul style={{ listStyle: 'none', padding: 0, margin: 0 }}>
                {adjuntos.map((a, i) => (
                  <li key={i} style={{ padding: '4px 0', fontSize: 14 }}>
                    <a href={a.url} target="_blank" rel="noopener noreferrer">{a.name}</a>
                  </li>
                ))}
              </ul>
            </div>
          )}

          {/* Reapertura section */}
          {ticket.reaperturado && ticket.motivoReapertura && (
            <div style={{ borderTop: '1px solid #dee2e6', paddingTop: 12 }}>
              <h5 style={{ fontSize: 16 }}>Reaperturado</h5>
              <TextField label="Motivo" value={ticket.motivoReapertura} multiline rows={3} readOnly />
              <Separator />
            </div>
          )}

          {/* Comments history */}
          {comments.length > 0 && (
            <div>
              <Label>Comentarios</Label>
              <Stack tokens={{ childrenGap: 8 }}>
                {comments.map(c => (
                  <div key={c.id} style={{ background: '#f8f9fa', borderRadius: 4, padding: '8px 12px', fontSize: 13 }}>
                    <strong>{c.author}</strong>
                    <span style={{ color: '#6c757d', marginLeft: 8, fontSize: 12 }}>{c.date}</span>
                    <p style={{ margin: '4px 0 0' }}>{c.text}</p>
                    {c.adjuntos.length > 0 && (
                      <ul style={{ listStyle: 'none', padding: 0, margin: '6px 0 0' }}>
                        {c.adjuntos.map((a, ai) => (
                          <li key={ai} style={{ fontSize: 12 }}>
                            <a href={a.url} target="_blank" rel="noopener noreferrer">📎 {a.name}</a>
                          </li>
                        ))}
                      </ul>
                    )}
                  </div>
                ))}
              </Stack>
            </div>
          )}

          {/* Comment / action area — visible when user has a role */}
          {(iamApproval || isProcessManager) && !isClosed && (
            <div>
              <TextField
                label="Comentarios"
                multiline
                rows={3}
                value={comentario}
                onChange={(_, v) => setComentario(v ?? '')}
                required={!iamApproval || (iamApproval && !isProcessManager)}
              />
              <div style={{ marginTop: 8 }}>
                <input
                  type="file"
                  ref={commentFileRef}
                  style={{ display: 'none' }}
                  onChange={e => setSelectedCommentFile(e.target.files?.[0] ?? null)}
                />
                <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }}>
                  <PrimaryButton
                    iconProps={{ iconName: 'Send' }}
                    text="Publicar comentario"
                    onClick={handlePublishComment}
                    disabled={isSubmitting || (!comentario.trim() && !selectedCommentFile)}
                  />
                  <DefaultButton
                    iconProps={{ iconName: 'Attach' }}
                    text="Cargar archivo"
                    onClick={() => commentFileRef.current?.click()}
                    disabled={isSubmitting}
                  />
                  {selectedCommentFile && (
                    <span style={{ fontSize: 13 }}>
                      {selectedCommentFile.name}{' '}
                      <IconButton iconProps={{ iconName: 'Cancel' }} title="Eliminar" onClick={() => setSelectedCommentFile(null)} />
                    </span>
                  )}
                  
                </Stack>
              </div>
            </div>
          )}

          {/* Register KB + Solución detallada (process manager closing) */}
          {isProcessManager && ticket.status === 'Assigned' && (
            <Stack tokens={stackTokens}>
              <Dropdown
                label="¿Deseas guardar en la base de conocimientos?"
                options={kbOptions}
                selectedKey={registrarKB}
                onChange={(_, o) => setRegistrarKB(o?.key as string)}
              />
              {registrarKB === 'Sí' && (
                <TextField
                  label="Solución detallada"
                  multiline
                  rows={4}
                  value={solucionDetallada}
                  onChange={(_, v) => setSolucionDetallada(v ?? '')}
                />
              )}
            </Stack>
          )}

          {/* Closed — read-only solution */}
          {isClosed && ticket.solucionDetallada && (
            <TextField label="Solución detallada" multiline rows={4} value={ticket.solucionDetallada} disabled />
          )}

          {/* Action buttons */}
          {(iamApproval || isProcessManager) && (
            <Stack horizontal tokens={{ childrenGap: 8 }} styles={{ root: { justifyContent: 'flex-end', paddingBottom: 24 } }}>
              {isProcessManager && ticket.status === 'Assigned' ? (
                <TooltipHost content={!ticket.atencion ? 'Debes marcar el ticket "En Atención" antes de cerrarlo' : ''}>
                  <PrimaryButton
                    text="Cerrar Ticket"
                    iconProps={{ iconName: 'CheckMark' }}
                    onClick={handleClose}
                    disabled={isSubmitting || !ticket.atencion}
                  />
                </TooltipHost>
              ) : (
                <>
                  <PrimaryButton text="Aprobar" onClick={handleApprove} disabled={isSubmitting} />
                  <DefaultButton
                    text="Rechazar"
                    onClick={handleReject}
                    disabled={isSubmitting}
                    styles={{ root: { background: '#dc3545', color: '#fff', borderColor: '#dc3545' } }}
                  />
                </>
              )}
              <DefaultButton text="Cerrar" onClick={onDismiss} />
              {isSubmitting && <Spinner size={SpinnerSize.small} />}
            </Stack>
          )}
        </Stack>
      )}
    </Panel>
  );
};

export default DetalleTicket;