import { LocationOnOutlinedIcon } from "@/components/sci/icons";

const fact = (value: string) => value || "—";

export function PersonCardFacts({
  rnokpp,
  birthDateWithAge,
  location,
  fullPosition,
  note,
  positionTitle,
  militaryId,
  arrivedFrom,
}: {
  rnokpp: string;
  birthDateWithAge: string;
  location: string;
  fullPosition: string;
  note: string;
  positionTitle: string;
  militaryId: string;
  arrivedFrom: string;
}) {
  return (
    <>
      <span>
        <strong>РНОКПП</strong>
        {fact(rnokpp)}
      </span>
      <span>
        <strong>Дата народження</strong>
        {fact(birthDateWithAge)}
      </span>
      <span className="person-location-highlight">
        <strong>
          <LocationOnOutlinedIcon fontSize="small" />
          Поточне місцеперебування
        </strong>
        <span className="person-location-value">{location || "Не вказано"}</span>
      </span>
      <div className="person-position-note-row">
        <span className="person-full-position-widget">
          <strong>Повна посада</strong>
          <span className="person-full-position-value">{fact(fullPosition)}</span>
        </span>
        <span className="person-note-widget">
          <strong>Примітка</strong>
          <span className="person-note-value">{fact(note)}</span>
        </span>
      </div>
      <span>
        <strong>Посада</strong>
        {fact(positionTitle)}
      </span>
      <span>
        <strong>Військовий квиток</strong>
        {fact(militaryId)}
      </span>
      <span>
        <strong>Звідки прибув</strong>
        {fact(arrivedFrom)}
      </span>
    </>
  );
}
