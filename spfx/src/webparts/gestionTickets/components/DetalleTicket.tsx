import * as React from 'react';
import { Panel, PanelType } from '@fluentui/react';

export interface IDetalleTicketProps {
  isOpen: boolean;
  onDismiss: () => void;
}

const DetalleTicket: React.FC<IDetalleTicketProps> = ({ isOpen, onDismiss }) => {
  return (
    <Panel
      isOpen={isOpen}
      onDismiss={onDismiss}
      headerText="Detalle de Ticket"
      closeButtonAriaLabel="Cerrar"
      type={PanelType.large}
    >
      <p>Contenido del panel para el detalle del ticket.</p>
      {/* Aquí puedes agregar más contenido para el detalle del ticket */}
    </Panel>
  );
};

export default DetalleTicket;