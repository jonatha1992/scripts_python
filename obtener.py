import pandas as pd

def parse_vcard(vcf_file):
    contacts = []
    current_contact = {}
    
    with open(vcf_file, 'r', encoding='utf-8') as file:
        for line in file:
            line = line.strip()
            if line == 'BEGIN:VCARD':
                current_contact = {}
            elif line.startswith('FN:'):
                current_contact['Nombre'] = line[3:]
            elif line.startswith('EMAIL'):
                email = line.split(':')[1]
                current_contact['Email'] = email
            elif line == 'END:VCARD':
                if current_contact:
                    contacts.append(current_contact)
    
    return contacts

def main():
    # Leer archivo VCF
    vcf_file = 'contacts.vcf'
    contacts = parse_vcard(vcf_file)
    
    # Crear DataFrame
    df = pd.DataFrame(contacts)
    
    # Exportar a Excel
    excel_file = 'contacts.xlsx'
    df.to_excel(excel_file, index=False)
    print(f'Archivo Excel generado: {excel_file}')

if __name__ == '__main__':
    main()