// avoids "TS2304: Cannot find name 'CONSTANT'" TSC error on builds because these
// strings are not defined in the code, rather webpack replaces them on build

declare var Environment: string;
declare var HttpCreateTeam: string;