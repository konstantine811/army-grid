import { Chip } from "@/components/sci/SciPrimitives";
import { PersonCardName } from "./PersonCardName";

export function PersonCardIdentity({
  name,
  callSign,
  rosterStatus,
  rank,
  positionIndex,
  serviceType,
}: {
  name: string;
  callSign: string;
  rosterStatus: string;
  rank: string;
  positionIndex: string;
  serviceType: string;
}) {
  return (
    <div className="person-card-identity">
      {callSign || rosterStatus ? (
        <div className="person-callsign-row">
          {callSign ? (
            <span className="person-callsign" title="Позивний">
              <span className="person-callsign-label">позивний</span>
              <strong>{callSign}</strong>
            </span>
          ) : null}
          {rosterStatus ? (
            <span className="person-roster-status" title="Статус">
              <span className="person-callsign-label">статус</span>
              <strong>{rosterStatus}</strong>
            </span>
          ) : null}
        </div>
      ) : null}
      <PersonCardName name={name} />
      <div className="person-action-tags">
        {rank ? <Chip label={rank} size="small" /> : null}
        {positionIndex ? (
          <Chip label={`Посада: ${positionIndex}`} size="small" />
        ) : null}
        {serviceType ? <Chip label={serviceType} size="small" /> : null}
      </div>
    </div>
  );
}
