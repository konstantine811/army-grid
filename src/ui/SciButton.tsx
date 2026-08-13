import type { ButtonHTMLAttributes, LabelHTMLAttributes, ReactNode } from "react";
import { cn } from "./cn";
import "./sci-ui.css";

type SciButtonVariant = "EXEC" | "OUTLINE" | "GHOST" | "ABORT";
type SciButtonSize = "SM" | "MD" | "LG";

type CommonProps = {
  variant?: SciButtonVariant;
  size?: SciButtonSize;
  startIcon?: ReactNode;
  endIcon?: ReactNode;
};

type ButtonProps = CommonProps &
  ButtonHTMLAttributes<HTMLButtonElement> & {
    as?: "button";
  };

type LabelProps = CommonProps &
  LabelHTMLAttributes<HTMLLabelElement> & {
    as: "label";
    disabled?: boolean;
  };

export type SciButtonProps = ButtonProps | LabelProps;

export function SciButton(props: SciButtonProps) {
  const {
    as = "button",
    children,
    className,
    disabled = false,
    endIcon,
    size = "MD",
    startIcon,
    variant = "OUTLINE",
    ...rest
  } = props;
  const content = (
    <>
      {startIcon && <span className="sci-button-icon">{startIcon}</span>}
      <span className="sci-button-label">{children}</span>
      {endIcon && <span className="sci-button-icon">{endIcon}</span>}
    </>
  );
  const classes = cn(
    "sci-button",
    `sci-button--${variant.toLowerCase()}`,
    `sci-button--${size.toLowerCase()}`,
    disabled && "sci-button--disabled",
    className,
  );

  if (as === "label") {
    return (
      <label
        className={classes}
        aria-disabled={disabled || undefined}
        {...(rest as LabelHTMLAttributes<HTMLLabelElement>)}
      >
        {content}
      </label>
    );
  }

  return (
    <button
      className={classes}
      disabled={disabled}
      type="button"
      {...(rest as ButtonHTMLAttributes<HTMLButtonElement>)}
    >
      {content}
    </button>
  );
}
