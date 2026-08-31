import {
  Children,
  createElement,
  type CSSProperties,
  type ChangeEvent,
  type ElementType,
  type HTMLAttributes,
  type ReactNode,
} from "react";
import { Button as SciButton } from "@/components/ui/button/button";
import { Input } from "@/components/ui/input/input";
import {
  Select as SciSelect,
  SelectContent,
  SelectItem,
  SelectTrigger,
  SelectValue,
} from "@/components/ui/select/select";
import {
  Dialog as SciDialog,
  DialogBody,
  DialogContent,
  DialogFooter,
  DialogHeader,
  DialogTitle as SciDialogTitle,
} from "@/components/ui/dialog/dialog";
import { Alert as SciAlert } from "@/components/ui/alert/alert";
import { Badge } from "@/components/ui/badge/badge";
import { Progress } from "@/components/ui/progress/progress";
import { Switch as SciSwitch } from "@/components/ui/switch/switch";
import { Checkbox as SciCheckbox } from "@/components/ui/checkbox/checkbox";
import { Textarea } from "@/components/ui/textarea/textarea";
import { cn } from "@/lib/utils";

type Sx = CSSProperties & Record<string, unknown>;

const spacingUnit = 8;

function sxToStyle(sx?: Sx): CSSProperties | undefined {
  if (!sx) return undefined;

  const style: CSSProperties = {};
  Object.entries(sx).forEach(([key, value]) => {
    if (value == null || typeof value === "object") return;
    const nextValue =
      typeof value === "number" &&
      ["m", "mt", "mr", "mb", "ml", "mx", "my", "p", "pt", "pr", "pb", "pl", "px", "py", "gap"].includes(key)
        ? value * spacingUnit
        : value;

    switch (key) {
      case "m":
        style.margin = nextValue as CSSProperties["margin"];
        break;
      case "mt":
        style.marginTop = nextValue as CSSProperties["marginTop"];
        break;
      case "mr":
        style.marginRight = nextValue as CSSProperties["marginRight"];
        break;
      case "mb":
        style.marginBottom = nextValue as CSSProperties["marginBottom"];
        break;
      case "ml":
        style.marginLeft = nextValue as CSSProperties["marginLeft"];
        break;
      case "mx":
        style.marginLeft = nextValue as CSSProperties["marginLeft"];
        style.marginRight = nextValue as CSSProperties["marginRight"];
        break;
      case "my":
        style.marginTop = nextValue as CSSProperties["marginTop"];
        style.marginBottom = nextValue as CSSProperties["marginBottom"];
        break;
      case "p":
        style.padding = nextValue as CSSProperties["padding"];
        break;
      case "pt":
        style.paddingTop = nextValue as CSSProperties["paddingTop"];
        break;
      case "pr":
        style.paddingRight = nextValue as CSSProperties["paddingRight"];
        break;
      case "pb":
        style.paddingBottom = nextValue as CSSProperties["paddingBottom"];
        break;
      case "pl":
        style.paddingLeft = nextValue as CSSProperties["paddingLeft"];
        break;
      case "px":
        style.paddingLeft = nextValue as CSSProperties["paddingLeft"];
        style.paddingRight = nextValue as CSSProperties["paddingRight"];
        break;
      case "py":
        style.paddingTop = nextValue as CSSProperties["paddingTop"];
        style.paddingBottom = nextValue as CSSProperties["paddingBottom"];
        break;
      default:
        (style as Record<string, unknown>)[key] = nextValue;
    }
  });
  return style;
}

type BoxProps = HTMLAttributes<HTMLElement> & {
  component?: ElementType;
  sx?: Sx;
};

export function Box({ component, sx, style, ...props }: BoxProps) {
  return createElement(component ?? "div", {
    style: { ...sxToStyle(sx), ...style },
    ...props,
  });
}

type StackProps = HTMLAttributes<HTMLDivElement> & {
  direction?: "row" | "column" | Record<string, "row" | "column">;
  spacing?: number;
  justifyContent?: string | Record<string, string>;
  alignItems?: string | Record<string, string>;
  flexWrap?: React.CSSProperties["flexWrap"];
  sx?: Sx;
};

export function Stack({
  direction = "column",
  spacing = 0,
  justifyContent,
  alignItems,
  flexWrap,
  sx,
  style,
  ...props
}: StackProps) {
  const resolvedDirection =
    typeof direction === "string" ? direction : direction.md ?? direction.xs ?? "column";
  const resolveResponsive = (value?: string | Record<string, string>) => {
    if (!value) return undefined;
    if (typeof value === "string") return value;
    return value.md ?? value.xs ?? Object.values(value)[0];
  };
  return (
    <div
      style={{
        display: "flex",
        flexDirection: resolvedDirection,
        gap: spacing * spacingUnit,
        justifyContent: resolveResponsive(justifyContent),
        alignItems: resolveResponsive(alignItems),
        flexWrap,
        ...sxToStyle(sx),
        ...style,
      }}
      {...props}
    />
  );
}

type TypographyProps = HTMLAttributes<HTMLElement> & {
  component?: ElementType;
  variant?: string;
  color?: string;
  noWrap?: boolean;
  sx?: Sx;
};

export function Typography({
  component,
  variant,
  color,
  noWrap,
  sx,
  className,
  style,
  ...props
}: TypographyProps) {
  const Component = component ?? (variant?.startsWith("h") ? variant : "p");
  const resolvedColor =
    color === "text.secondary"
      ? "var(--text-muted)"
      : color?.includes(".")
        ? undefined
        : color;
  return (
    createElement(Component, {
      className: cn("sci-text", variant && `sci-text-${variant}`, className),
      style: {
        color: resolvedColor,
        overflow: noWrap ? "hidden" : undefined,
        textOverflow: noWrap ? "ellipsis" : undefined,
        whiteSpace: noWrap ? "nowrap" : undefined,
        ...sxToStyle(sx),
        ...style,
      },
      ...props,
    })
  );
}

type ButtonProps = Omit<React.ComponentProps<typeof SciButton>, "variant" | "size"> & {
  variant?: "contained" | "outlined" | "text" | "EXEC" | "OUTLINE" | "GHOST" | "ABORT";
  startIcon?: ReactNode;
  endIcon?: ReactNode;
  component?: ElementType;
  href?: string;
  target?: string;
  rel?: string;
  fullWidth?: boolean;
  size?: "small" | "medium" | "large" | "SM" | "MD" | "LG";
  sx?: Sx;
};

export function Button({
  variant = "OUTLINE",
  startIcon,
  endIcon,
  component,
  fullWidth,
  size,
  children,
  sx,
  style,
  ...props
}: ButtonProps) {
  const sciVariant =
    variant === "contained" || variant === "EXEC"
      ? "EXEC"
      : variant === "text" || variant === "GHOST"
        ? "GHOST"
        : variant === "ABORT"
          ? "ABORT"
          : "OUTLINE";
  const sciSize = size === "small" || size === "SM" ? "SM" : size === "large" || size === "LG" ? "LG" : "MD";
  const sxStyle = sxToStyle(sx);
  if (variant === "contained" || variant === "EXEC") {
    delete sxStyle?.color;
  }
  const buttonStyle = {
    width: fullWidth ? "100%" : undefined,
    ...sxStyle,
    ...style,
  };

  if (component) {
    return (
      <SciButton asChild variant={sciVariant} size={sciSize} style={buttonStyle}>
        {createElement(component, props, <>
          {startIcon}
          {children}
          {endIcon}
        </>)}
      </SciButton>
    );
  }

  return (
    <SciButton variant={sciVariant} size={sciSize} style={buttonStyle} {...props}>
      {startIcon}
      {children}
      {endIcon}
    </SciButton>
  );
}

type IconButtonProps = React.ButtonHTMLAttributes<HTMLButtonElement> & {
  size?: "small" | "medium" | "large";
  sx?: Sx;
};

export function IconButton({ className, size = "medium", sx, style, ...props }: IconButtonProps) {
  return (
    <SciButton
      className={cn("sci-icon-button", className)}
      size={size === "large" ? "LG" : "SM"}
      variant="GHOST"
      style={{ width: size === "large" ? 44 : 34, padding: 0, ...sxToStyle(sx), ...style }}
      {...props}
    />
  );
}

type AlertProps = HTMLAttributes<HTMLDivElement> & {
  severity?: "info" | "success" | "warning" | "error";
  variant?: string;
  icon?: ReactNode;
  action?: ReactNode;
  sx?: Sx;
};

export function Alert({ severity = "info", variant: _variant, sx, style, className, ...props }: AlertProps) {
  const { icon: _icon, action, children, ...restProps } = props;
  const alertVariant: "STATUS" | "WARNING" | "CRITICAL" | "INFO" =
    severity === "success"
      ? "STATUS"
      : severity === "warning"
        ? "WARNING"
        : severity === "error"
          ? "CRITICAL"
          : "INFO";

  return (
    <SciAlert
      variant={alertVariant}
      className={cn("sci-alert", `sci-alert-${severity}`, className)}
      style={{ ...sxToStyle(sx), ...style }}
      {...restProps}
    >
      {children}
      {action ? <span className="sci-alert-action">{action}</span> : null}
    </SciAlert>
  );
}

export function LinearProgress({ sx, style }: { color?: string; sx?: Sx; style?: CSSProperties }) {
  return (
    <div className="sci-linear-progress" style={{ ...sxToStyle(sx), ...style }}>
      <Progress indeterminate label="PROCESSING" showValue={false} />
    </div>
  );
}

type TextFieldProps = Omit<React.InputHTMLAttributes<HTMLInputElement>, "size" | "onChange"> & {
  label?: string;
  select?: boolean;
  multiline?: boolean;
  rows?: number;
  size?: "small" | "medium";
  fullWidth?: boolean;
  children?: ReactNode;
  sx?: Sx;
  slotProps?: Record<string, unknown>;
  suffix?: ReactNode;
  onChange?: (event: ChangeEvent<HTMLInputElement | HTMLSelectElement>) => void;
};

export function TextField({
  label,
  select,
  multiline,
  rows,
  children,
  sx,
  style,
  className,
  onChange,
  size: _size,
  fullWidth: _fullWidth,
  slotProps: _slotProps,
  suffix,
  ...props
}: TextFieldProps) {
  if (select) {
    const currentValue = String(props.value ?? "");
    return (
      <label className={cn("sci-field", className)} style={{ ...sxToStyle(sx), ...style }}>
        {label && <span>{label}</span>}
        <Select
          value={currentValue}
          disabled={props.disabled}
          onChange={(event) => onChange?.(event as unknown as ChangeEvent<HTMLSelectElement>)}
        >
          {children}
        </Select>
      </label>
    );
  }

  if (multiline) {
    return (
      <Textarea
        label={label}
        className={className}
        rows={rows}
        style={{ ...sxToStyle(sx), ...style }}
        onChange={(event) =>
          onChange?.(event as unknown as ChangeEvent<HTMLInputElement | HTMLSelectElement>)
        }
        value={props.value as string | number | readonly string[] | undefined}
        placeholder={props.placeholder}
        disabled={props.disabled}
        name={props.name}
        id={props.id}
      />
    );
  }

  return (
    <Input
      label={label}
      className={className}
      style={{ ...sxToStyle(sx), ...style }}
      onChange={onChange as React.ChangeEventHandler<HTMLInputElement>}
      suffix={suffix}
      {...props}
    />
  );
}

export function MenuItem({
  value: _value,
  children: _children,
  disabled: _disabled,
  className: _className,
}: {
  value?: string | number;
  children?: ReactNode;
  disabled?: boolean;
  className?: string;
}) {
  return null;
}

type SelectProps = {
  value?: string;
  onChange?: (event: { target: { value: string } }) => void;
  children?: ReactNode;
  displayEmpty?: boolean;
  disabled?: boolean;
  className?: string;
  sx?: Sx;
  labelId?: string;
  size?: string;
  renderValue?: (value: string) => ReactNode;
  MenuProps?: unknown;
};

export function Select({ value, onChange, children, disabled, className, sx, renderValue }: SelectProps) {
  const options = Children.toArray(children);
  return (
    <SciSelect value={value} onValueChange={(nextValue) => onChange?.({ target: { value: nextValue } })} disabled={disabled}>
      <SelectTrigger className={className} style={sxToStyle(sx)}>
        <SelectValue>{renderValue && value ? renderValue(value) : undefined}</SelectValue>
      </SelectTrigger>
      <SelectContent>
        {options.map((child, index) => {
          if (!child || typeof child !== "object" || !("props" in child)) return child;
          const item = child as {
            props: {
              value?: string;
              children?: ReactNode;
              disabled?: boolean;
              className?: string;
            };
          };
          return (
            <SelectItem
              key={`${item.props.value ?? index}`}
              value={String(item.props.value ?? "")}
              disabled={item.props.disabled}
              className={item.props.className}
            >
              {item.props.children}
            </SelectItem>
          );
        })}
      </SelectContent>
    </SciSelect>
  );
}

export function Divider({
  sx,
  style,
  orientation,
  flexItem,
}: {
  sx?: Sx;
  style?: CSSProperties;
  orientation?: "horizontal" | "vertical" | string;
  flexItem?: boolean;
}) {
  return (
    <div
      className={cn(
        "sci-divider",
        orientation === "vertical" && "sci-divider-vertical",
        flexItem && "sci-divider-flex-item",
      )}
      style={{ ...sxToStyle(sx), ...style }}
    />
  );
}

export function Switch({
  checked,
  defaultChecked,
  onChange,
  disabled,
  onClick,
  onMouseDown,
}: {
  checked?: boolean;
  defaultChecked?: boolean;
  onChange?: (event: ChangeEvent<HTMLInputElement>, checked: boolean) => void;
  color?: string;
  disabled?: boolean;
  size?: string;
  onClick?: React.MouseEventHandler;
  onMouseDown?: React.MouseEventHandler;
  slotProps?: unknown;
}) {
  return (
    <span onClick={onClick} onMouseDown={onMouseDown}>
      <SciSwitch
        checked={checked}
        defaultChecked={defaultChecked}
        disabled={disabled}
        onCheckedChange={(nextChecked) =>
          onChange?.(
            { target: { checked: nextChecked } } as ChangeEvent<HTMLInputElement>,
            nextChecked,
          )
        }
      />
    </span>
  );
}

export function Checkbox({
  checked,
  defaultChecked,
  onCheckedChange,
  disabled,
  label,
  id,
  className,
}: {
  checked?: boolean | "indeterminate";
  defaultChecked?: boolean;
  onCheckedChange?: (checked: boolean | "indeterminate") => void;
  disabled?: boolean;
  label?: string;
  id?: string;
  className?: string;
}) {
  return (
    <SciCheckbox
      id={id}
      label={label}
      className={className}
      checked={checked}
      defaultChecked={defaultChecked}
      disabled={disabled}
      onCheckedChange={onCheckedChange}
    />
  );
}

export function Chip({
  label,
  className,
  color,
  component,
  sx,
  style,
}: {
  label?: ReactNode;
  className?: string;
  size?: string;
  color?: string;
  variant?: string;
  component?: ElementType;
  sx?: Sx;
  style?: CSSProperties;
}) {
  const badgeVariant =
    color === "warning"
      ? "WARNING"
      : color === "error"
        ? "CRITICAL"
        : color === "default"
        ? "OFFLINE"
        : "ACTIVE";

  if (component) {
    return createElement(component, {
      className,
      style: { ...sxToStyle(sx), ...style },
      children: label,
    });
  }

  return (
    <Badge variant={badgeVariant} className={className} style={{ ...sxToStyle(sx), ...style }}>
      {label}
    </Badge>
  );
}

type DialogSlotProps = {
  paper?: {
    className?: string;
  };
};

export function Dialog({
  open,
  onClose,
  children,
  slotProps,
  modal = true,
}: {
  open?: boolean;
  onClose?: () => void;
  children?: ReactNode;
  maxWidth?: string;
  fullWidth?: boolean;
  modal?: boolean;
  slotProps?: DialogSlotProps;
}) {
  return (
    <SciDialog
      open={open}
      modal={modal}
      onOpenChange={(nextOpen) => !nextOpen && onClose?.()}
    >
      <DialogContent
        hideOverlay={!modal}
        className={slotProps?.paper?.className}
      >
        {children}
      </DialogContent>
    </SciDialog>
  );
}

export function DialogTitle({ children }: { children?: ReactNode }) {
  return (
    <DialogHeader>
      <SciDialogTitle>{children}</SciDialogTitle>
    </DialogHeader>
  );
}

export function DialogContentBox({ children, sx, style }: { children?: ReactNode; sx?: Sx; style?: CSSProperties }) {
  return <DialogBody style={{ ...sxToStyle(sx), ...style }}>{children}</DialogBody>;
}

export { DialogContentBox as DialogContent };

export function DialogActions({ children }: { children?: ReactNode }) {
  return <DialogFooter>{children}</DialogFooter>;
}
