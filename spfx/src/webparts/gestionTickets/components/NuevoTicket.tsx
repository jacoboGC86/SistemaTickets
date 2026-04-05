import * as React from 'react';
import {
  Panel,
  PanelType,
  Stack,
  TextField,
  Dropdown,
  IDropdownOption,
  Checkbox,
  PrimaryButton,
  DefaultButton,
  IStackTokens,
  Image,
  ImageFit,
} from '@fluentui/react';

export interface INuevoTicketProps {
  isOpen: boolean;
  onDismiss: () => void;
}

const typeOptions: IDropdownOption[] = [
  { key: 'Request', text: 'Solicitud' },
  { key: 'Incident', text: 'Incidente' },
  { key: 'Change', text: 'Cambio' },
];

const categoryOptions: IDropdownOption[] = [
  { key: 'none', text: 'Seleccione una opción' },
  { key: 'cat1', text: 'Categoría 1' },
  { key: 'cat2', text: 'Categoría 2' },
];

const departmentOptions: IDropdownOption[] = [
  { key: 'dept1', text: 'Departamento 1' },
  { key: 'dept2', text: 'Departamento 2' },
];

const plantOptions: IDropdownOption[] = [
  { key: 'Aguascalientes', text: 'Aguascalientes' },
  { key: 'Puebla', text: 'Puebla' },
];

const urgencyOptions: IDropdownOption[] = [
  { key: 'Alta', text: 'Alta' },
  { key: 'Media', text: 'Media' },
  { key: 'Baja', text: 'Baja' },
];

const stackTokens: IStackTokens = {
  childrenGap: 16,
};

const NuevoTicket: React.FC<INuevoTicketProps> = ({ isOpen, onDismiss }) => {
  const [ticketType, setTicketType] = React.useState<string>('');
  const [category, setCategory] = React.useState<string>('none');
  const [title, setTitle] = React.useState<string>('');
  const [description, setDescription] = React.useState<string>('');
  const [department, setDepartment] = React.useState<string>('');
  const [plant, setPlant] = React.useState<string>('Aguascalientes');
  const [urgency, setUrgency] = React.useState<string>('');
  const [registerForOther, setRegisterForOther] = React.useState<boolean>(false);
  const [securityImpact, setSecurityImpact] = React.useState<string>('No');
  const [securityDescription, setSecurityDescription] = React.useState<string>('');
  const [attachments, setAttachments] = React.useState<File[]>([]);
  const fileInputRef = React.useRef<HTMLInputElement | null>(null);

  const onAttachmentClick = () => {
    fileInputRef.current?.click();
  };

  const onFilesSelected = (ev: React.ChangeEvent<HTMLInputElement>) => {
    if (ev.target.files) {
      //setAttachments(Array.from(ev.target.files));
    }
  };

  return (
    <Panel
      isOpen={isOpen}
      onDismiss={onDismiss}
      headerText="Nuevo Ticket"
      closeButtonAriaLabel="Cerrar"
      type={PanelType.large}
    >
      <Stack tokens={stackTokens}>
        <Stack horizontal wrap tokens={{ childrenGap: 12 }}>
          <TextField label="Fecha de registro" value={new Date().toLocaleDateString()} disabled styles={{ root: { minWidth: 260 } }} />
        </Stack>

        <Stack horizontal wrap tokens={{ childrenGap: 12 }}>
          <Dropdown
            label="Tipo"
            placeholder="Seleccione una opción"
            options={typeOptions}
            selectedKey={ticketType}
            onChange={(_, option) => setTicketType((option?.key as string) || '')}
            styles={{ root: { minWidth: 260 } }}
          />
          <Dropdown
            label="Categoría"
            options={categoryOptions}
            selectedKey={category}
            onChange={(_, option) => setCategory((option?.key as string) || 'none')}
            styles={{ root: { minWidth: 260 } }}
          />
        </Stack>

        <TextField
          label="Título"
          value={title}
          onChange={(_, newValue) => setTitle(newValue || '')}
          styles={{ root: { width: '100%' } }}
        />

        <TextField
          label="Descripción detallada"
          value={description}
          onChange={(_, newValue) => setDescription(newValue || '')}
          multiline
          rows={4}
          styles={{ root: { width: '100%' } }}
        />

        <Stack horizontal wrap tokens={{ childrenGap: 12 }}>
          <Dropdown
            label="Departamento"
            options={departmentOptions}
            selectedKey={department}
            onChange={(_, option) => setDepartment((option?.key as string) || '')}
            styles={{ root: { minWidth: 260 } }}
          />
          <Dropdown
            label="Planta"
            options={plantOptions}
            selectedKey={plant}
            onChange={(_, option) => setPlant((option?.key as string) || 'Aguascalientes')}
            styles={{ root: { minWidth: 260 } }}
          />
        </Stack>

        <Stack horizontal wrap tokens={{ childrenGap: 12 }}>
          <Dropdown
            label="Urgencia"
            options={urgencyOptions}
            selectedKey={urgency}
            onChange={(_, option) => setUrgency((option?.key as string) || '')}
            styles={{ root: { minWidth: 260 } }}
          />
          <Checkbox
            label="¿Estás registrando este ticket a nombre de otra persona?"
            checked={registerForOther}
            onChange={(_, checked) => setRegisterForOther(!!checked)}
            styles={{ root: { minWidth: 260 } }}
          />
        </Stack>

        {registerForOther && (
          <TextField
            label="A nombre de"
            placeholder="Nombre de la persona"
            styles={{ root: { width: '100%' } }}
          />
        )}

        {ticketType === 'Change' && (
          <Stack tokens={{ childrenGap: 12 }}>
            <Dropdown
              label="¿Este cambio impacta la seguridad de la información?"
              options={[
                { key: 'No', text: 'No' },
                { key: 'Sí', text: 'Sí' },
              ]}
              selectedKey={securityImpact}
              onChange={(_, option) => setSecurityImpact((option?.key as string) || 'No')}
              styles={{ root: { minWidth: 260 } }}
            />
            {securityImpact === 'Sí' && (
              <TextField
                label="Describe los impactos a la seguridad de la información"
                value={securityDescription}
                onChange={(_, newValue) => setSecurityDescription(newValue || '')}
                multiline
                rows={3}
                styles={{ root: { width: '100%' } }}
              />
            )}
          </Stack>
        )}

        <Stack tokens={{ childrenGap: 8 }}>
          <DefaultButton text="Cargar archivos" onClick={onAttachmentClick} />
          <input
            type="file"
            multiple
            hidden
            ref={fileInputRef}
            onChange={onFilesSelected}
          />
          {attachments.length > 0 && (
            <Stack tokens={{ childrenGap: 4 }}>
              {attachments.map((file, index) => (
                <span key={index}>{file.name}</span>
              ))}
            </Stack>
          )}
        </Stack>

        <Stack horizontal horizontalAlign="end" tokens={{ childrenGap: 12 }}>
          <PrimaryButton text="Enviar" onClick={() => undefined} />
          <DefaultButton text="Cancelar" onClick={onDismiss} />
        </Stack>
      </Stack>
    </Panel>
  );
};

export default NuevoTicket;
