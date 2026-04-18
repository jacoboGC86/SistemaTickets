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
} from '@fluentui/react';
import ListSvc from '../../../services/ListSvc';
import UserSvc from '../../../services/UserSvc';

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
    setComentario('');
    setMotivoReapertura('');
    setSolucionDetallada('');
    setRegistrarKB('No');
    setNewCategoryId(null);
    setSelectedCommentFile(null);
    setReassignQuery('');
    setReassignSuggestions([]);
    setSelectedReassign(null);
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
        `$filter=Id eq ${id}&$select=Id,Title,Status,TipoTicket,Descripcion,Department,Planta,Prioridad,TemplateConfiguracion,SolucionDetallada,Reaperturado,MotivoReapertura,Atencion,ImpactoSeguridadInformacion,DescripcionImpactoSeguridad,Created,Modified,CategoriaId,AuthorId,ProcessManagerId,ManagerId,ANombreDeId,Author/Title,Author/Id,Categoria/Title,ProcessManager/Title,ProcessManager/Id,Manager/Title,Manager/Id,ANombreDe/Title&$expand=Author,Categoria,ProcessManager,Manager,ANombreDe&$top=1`
      );
      const ticketItem = ticketItems?.[0];
      if (!ticketItem) throw new Error('Ticket no encontrado.');

      const currentUser = await UserSvc.GetCurrentUser();

      const uid: number = currentUser?.Id ?? 0;
      setCurrentUserId(uid);

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
        modified: ticketItem.Modified ?? '',
      };
      setTicket(td);
      setSolucionDetallada(td.solucionDetallada);

      const isPM = uid === td.processManagerId && td.status === 'Assigned';
      const isApproverStep = uid === td.managerId && td.status !== 'Assigned' && td.status !== 'Cerrado';
      setIsProcessManager(isPM);

      // Load entity type names for write operations
      const [ticketsType, aprobacionesType, comentariosType] = await Promise.all([
        ListSvc.getListItemEntityTypeFullName('Tickets'),
        ListSvc.getListItemEntityTypeFullName('Aprobaciones'),
        ListSvc.getListItemEntityTypeFullName('Comentarios').catch(() => 'SP.Data.ComentariosListItem'),
      ]);
      setEntityTypes({ tickets: ticketsType, aprobaciones: aprobacionesType, comentarios: comentariosType });

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
        `$select=Id,Title,Resultado,Responsable/Id,Responsable/Title,Responsable/EMail,Modified&$expand=Responsable&$filter=TicketId eq ${id}&$orderby=Id asc`
      );
      const rows: IAprobacionRow[] = (aprobItems || []).map((a: any) => ({
        id: a.Id,
        title: a.Title ?? '',
        resultado: a.Resultado ?? 'Pendiente',
        responsable: a.Responsable?.Title ?? null,
        responsableEmail: a.Responsable?.EMail ?? null,
        modified: a.Modified ?? null,
      }));
      setAprobaciones(rows);

      // Determine if current user is an approver
      const pendingRow = rows.find(r => r.resultado === 'Pendiente');
      if (pendingRow && td.status !== 'Assigned' && td.status !== 'Cerrado') {
        setIamApproval(isApproverStep || isPM);
      } else {
        setIamApproval(isPM);
      }

      // Load adjuntos from Expediente folder
      try {
        const relativeURL = ListSvc.getRelativeSiteURL();
        const filesResp = await ListSvc.getItems(
          'Expediente',
          undefined,
          `$select=FileLeafRef,ServerRelativeUrl&$filter=TicketId eq ${id}`
        ).catch(() => null);
        if (filesResp) {
          setAdjuntos((filesResp || []).map((f: any) => ({ name: f.FileLeafRef, url: f.ServerRelativeUrl })));
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
        setComments((commentItems || []).map((c: any) => ({
          id: c.Id,
          author: c.Author?.Title ?? '',
          text: c.Comentario ?? c.Title ?? '',
          date: c.Created ? new Date(c.Created).toLocaleString('es-MX') : '',
          adjuntos: [],
        })));
      } catch {
        // Comments list may not exist
      }

      dataLoaded.current = true;
    } catch (e) {
      console.error('Error loading ticket detail:', e);
      setErrorMessage('Error al cargar el detalle del ticket.');
    } finally {
      setIsLoading(false);
    }
  };

  const [entityTypes, setEntityTypes] = React.useState<{ tickets: string; aprobaciones: string; comentarios: string }>({ tickets: '', aprobaciones: '', comentarios: '' });

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

  const handleApprove = async (): Promise<void> => {
    if (!ticket) return;
    setErrorMessage('');
    setSubmitAction('approve');
    setIsSubmitting(true);
    try {
      const pendingRow = aprobaciones.find(r => r.resultado === 'Pendiente');
      if (pendingRow) {
        await ListSvc.putListItem('Aprobaciones', pendingRow.id, JSON.stringify({ '__metadata': { type: entityTypes.aprobaciones }, Resultado: 'Aprobado' }));
      }
      // Advance ticket status
      const tmpl = ticket.templateConfiguracion;
      if (tmpl) {
        const currentIdx = tmpl.approvalPath.findIndex(s => s.stepName === ticket.status);
        const nextStep = tmpl.approvalPath[currentIdx + 1];
        const newStatus = nextStep ? nextStep.stepName : 'Assigned';
        await ListSvc.putListItem('Tickets', ticket.id, JSON.stringify({ '__metadata': { type: entityTypes.tickets }, Status: newStatus }));
      }
      if (comentario.trim()) {
        await ListSvc.postListItem('Comentarios', JSON.stringify({ '__metadata': { type: entityTypes.comentarios }, Title: comentario, TicketId: ticket.id }));
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
      const updatePayload: Record<string, unknown> = { '__metadata': { type: entityTypes.tickets }, Status: 'Cerrado' };
      if (registrarKB === 'Sí' && solucionDetallada.trim()) {
        updatePayload.SolucionDetallada = solucionDetallada;
      }
      await ListSvc.putListItem('Tickets', ticket.id, JSON.stringify(updatePayload));
      if (comentario.trim()) {
        await ListSvc.postListItem('Comentarios', JSON.stringify({ '__metadata': { type: entityTypes.comentarios }, Title: comentario, TicketId: ticket.id }));
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
      await ListSvc.putListItem('Tickets', ticket.id, JSON.stringify({ '__metadata': { type: entityTypes.tickets }, CategoriaId: newCategoryId }));
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

  const handleUpComment = async (): Promise<void> => {
    if (!ticket || !comentario.trim()) { setErrorMessage('Ingrese un comentario antes de subir.'); return; }
    setIsSubmitting(true);
    try {
      await ListSvc.postListItem('Comentarios', JSON.stringify({ '__metadata': { type: entityTypes.comentarios }, Title: comentario, TicketId: ticket.id }));
      setComentario('');
      setSelectedCommentFile(null);
      setSuccessMessage('Comentario guardado.');
      dataLoaded.current = false;
      await loadAll(ticket.id);
    } catch (e) {
      console.error(e);
      setErrorMessage('Error al guardar el comentario.');
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
        steps.push({
          stepName: ap.stepName,
          resultado: row?.resultado ?? 'Pendiente',
          responsable: row?.responsable ?? ticket.managerName ?? null,
          responsableEmail: row?.responsableEmail ?? null,
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
                  <DefaultButton iconProps={{ iconName: 'Tag' }} text="Recategorizar" onClick={() => setShowRecategorizarBar(v => !v)} />
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
              <TextField label="Solicitante" value={ticket.solicitante} disabled />
            </Stack.Item>
            <Stack.Item grow={1} styles={{ root: { minWidth: 100 } }}>
              <TextField label="Folio" value={String(ticket.id)} disabled />
            </Stack.Item>
            <Stack.Item grow={2} styles={{ root: { minWidth: 150 } }}>
              <TextField label="Fecha de registro" value={ticket.fechaRegistro} disabled />
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
              <Dropdown
                label="Categoría"
                options={showRecategorizarBar ? categorias : [{ key: ticket.categoriaId ?? 0, text: ticket.categoriaTitle }]}
                selectedKey={showRecategorizarBar ? (newCategoryId ?? ticket.categoriaId) : ticket.categoriaId}
                disabled={!showRecategorizarBar}
                onChange={(_, o) => setNewCategoryId(o?.key as number)}
              />
            </Stack.Item>
          </Stack>

          {/* Recategorizar action bar */}
          {showRecategorizarBar && (
            <Stack horizontal tokens={{ childrenGap: 8 }} styles={{ root: { justifyContent: 'flex-end' } }}>
              <PrimaryButton text="Recategorizar" onClick={handleRecategorizar} disabled={isSubmitting} />
              <DefaultButton text="Cancelar" onClick={() => { setShowRecategorizarBar(false); setNewCategoryId(null); }} />
            </Stack>
          )}

          <TextField label="Título" value={ticket.title} disabled />

          <TextField label="Descripción detallada" value={ticket.descripcion} multiline rows={4} disabled />

          <Stack horizontal tokens={{ childrenGap: 12 }} styles={{ root: { flexWrap: 'wrap' } }}>
            <Stack.Item grow={1} styles={{ root: { minWidth: 140 } }}>
              <TextField label="Departamento" value={ticket.departamento} disabled />
            </Stack.Item>
            <Stack.Item grow={1} styles={{ root: { minWidth: 140 } }}>
              <TextField label="Planta" value={ticket.planta} disabled />
            </Stack.Item>
            {ticket.aNombreDe && (
              <Stack.Item grow={1} styles={{ root: { minWidth: 140 } }}>
                <TextField label="A nombre de" value={ticket.aNombreDe} disabled />
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
                  disabled
                />
              )}
            </Stack>
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
              <TextField label="Motivo" value={ticket.motivoReapertura} multiline rows={3} disabled />
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
                  <DefaultButton
                    iconProps={{ iconName: 'Attach' }}
                    text="Cargar archivo"
                    onClick={() => commentFileRef.current?.click()}
                  />
                  {selectedCommentFile && (
                    <span style={{ fontSize: 13 }}>
                      {selectedCommentFile.name}{' '}
                      <IconButton iconProps={{ iconName: 'Cancel' }} title="Eliminar" onClick={() => setSelectedCommentFile(null)} />
                    </span>
                  )}
                  <PrimaryButton text="Subir" onClick={handleUpComment} disabled={isSubmitting} />
                </Stack>
              </div>
            </div>
          )}

          {/* Register KB + Detailed solution (process manager closing) */}
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
                  label="Detailed solution"
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
            <TextField label="Detailed solution" multiline rows={4} value={ticket.solucionDetallada} disabled />
          )}

          {/* Action buttons */}
          {(iamApproval || isProcessManager) && (
            <Stack horizontal tokens={{ childrenGap: 8 }} styles={{ root: { justifyContent: 'flex-end', paddingBottom: 24 } }}>
              {isProcessManager && ticket.status === 'Assigned' ? (
                <PrimaryButton
                  text="Cerrar Ticket"
                  iconProps={{ iconName: 'CheckMark' }}
                  onClick={handleClose}
                  disabled={isSubmitting}
                />
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